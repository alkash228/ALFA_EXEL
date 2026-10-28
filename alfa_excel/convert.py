"""Поток TXT → XLSX.

Файл выгрузки читается с диска по строкам. В памяти одна строка и одна
пачка ячеек. Книга пишется в режиме constant_memory: строка уходит на диск
и не копится в списке.
"""

from __future__ import annotations

import datetime as dt
import io
import os
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterator

import xlsxwriter


EXCEL_MAX_DATA_ROWS = 1_048_575
SAMPLE_BYTES = 1024 * 1024
PROGRESS_EVERY = 4000

SKIPPED = object()


class ExportError(Exception):
    """Ошибка, которую можно показать пользователю."""


class Cancelled(ExportError):
    """Остановка по кнопке. Уже закрытая книга остаётся на диске."""


@dataclass
class Settings:
    encoding: str = "auto"
    mode: str = "auto"
    delimiter: str = "auto"
    qualifier: str = ""
    header: str = "auto"
    fmt_path: str = ""
    rows_per_file: int = 1_000_000
    max_line_chars: int = 2_000_000


@dataclass
class Layout:
    encoding: str
    mode: str
    delimiter: str = "\t"
    qualifier: str = ""
    header_from_file: bool = False
    synthetic_header: bool = False
    column_names: list[str] = field(default_factory=list)
    widths: list[int] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    preview_header: list[str] = field(default_factory=list)
    sample: list[list[str]] = field(default_factory=list)
    allow_multiline_quotes: bool = False
    locked_columns: int = 0
    drop_leading_blank: bool = False


@dataclass
class Progress:
    stage: str
    fraction: float = 0.0
    message: str = ""
    rows: int = 0
    bytes_read: int = 0
    file_size: int = 0
    files_done: int = 0
    rows_per_sec: float = 0.0


@dataclass
class ConvertResult:
    rows: int
    files: list[str]
    elapsed: float
    truncated_cells: int
    skipped_long: int
    ragged_rows: int
    encoding: str
    mode: str
    columns: int


@dataclass
class Stats:
    skipped_long: int = 0
    truncated: int = 0
    ragged: int = 0


class CountingRaw(io.RawIOBase):
    """Считает байты, которые буфер забирает с диска. Сам файл не копит."""

    def __init__(self, raw: io.BufferedIOBase) -> None:
        super().__init__()
        self._raw = raw
        self.read_bytes = 0

    def readable(self) -> bool:
        return True

    def readinto(self, buffer) -> int:
        data = self._raw.read(len(buffer))
        if not data:
            return 0
        count = len(data)
        buffer[:count] = data
        self.read_bytes += count
        return count

    def read(self, size: int = -1):
        data = self._raw.read(size)
        if data:
            self.read_bytes += len(data)
        return data

    def close(self) -> None:
        if not self._raw.closed:
            self._raw.close()
        super().close()


class Source:
    def __init__(self, path: str, encoding: str) -> None:
        self.size = os.path.getsize(path)
        try:
            raw = open(path, "rb", buffering=0)
        except OSError as exc:
            raise ExportError(f"Не удалось открыть файл: {exc}") from exc
        self.counter = CountingRaw(raw)
        buffered = io.BufferedReader(self.counter, buffer_size=1024 * 1024)
        self.text = io.TextIOWrapper(
            buffered,
            encoding=encoding,
            errors="replace",
            newline="",
        )

    def close(self) -> None:
        self.text.close()


def describe_delimiter(value: str) -> str:
    names = {
        "\t": "табуляция",
        ",": "запятая",
        ";": "точка с запятой",
        "|": "вертикальная черта",
        "\u00a0": "неразрывный пробел",
        "  ": "несколько пробелов",
        "\r\n": "CRLF",
        "\n": "LF",
        "\r": "CR",
    }
    return names.get(value, repr(value))


def unescape_terminator(raw: str) -> str:
    if raw == "":
        raise ExportError("Разделитель не может быть пустым.")
    pieces: list[str] = []
    index = 0
    while index < len(raw):
        if raw[index] == "\\" and index + 1 < len(raw):
            mapped = {"t": "\t", "n": "\n", "r": "\r", "0": "\0", "\\": "\\"}.get(raw[index + 1])
            if mapped is None:
                raise ExportError(
                    "В разделителе неизвестная последовательность. "
                    "Нужен обычный символ, табуляция, перевод строки или возврат каретки."
                )
            pieces.append(mapped)
            index += 2
            continue
        pieces.append(raw[index])
        index += 1
    text = "".join(pieces)
    if text == "":
        raise ExportError("Разделитель не может быть пустым.")
    if len(text) > 10:
        raise ExportError("Разделитель длиннее 10 символов.")
    return text


def resolve_encoding(path: str, choice: str) -> str:
    with open(path, "rb") as handle:
        head = handle.read(4)
        sample = head + handle.read(SAMPLE_BYTES)
    if choice != "auto":
        if choice == "utf-8" and head.startswith(b"\xef\xbb\xbf"):
            return "utf-8-sig"
        return choice
    if head.startswith(b"\xff\xfe") or head.startswith(b"\xfe\xff"):
        return "utf-16"
    if head.startswith(b"\xef\xbb\xbf"):
        return "utf-8-sig"
    if sample.count(b"\x00") > max(8, len(sample) // 4):
        return "utf-16-le"
    try:
        sample.decode("utf-8")
        return "utf-8"
    except UnicodeDecodeError as exc:
        if exc.start >= len(sample) - 4:
            return "utf-8"
    best = "cp1251"
    best_score = -1
    for encoding in ("cp1251", "cp866", "cp1252"):
        try:
            text = sample.decode(encoding)
        except UnicodeDecodeError:
            continue
        score = sum(1 for char in text if "\u0400" <= char <= "\u04ff")
        if score > best_score:
            best = encoding
            best_score = score
    return best if best_score > 0 else "latin-1"


def read_logical_line(text_file: io.TextIOBase, max_chars: int, stats: Stats):
    parts: list[str] = []
    size = 0
    while True:
        chunk = text_file.readline(256 * 1024)
        if chunk == "":
            if not parts:
                return None
            break
        size += len(chunk)
        if size > max_chars and not (chunk.endswith("\n") or chunk.endswith("\r")):
            while True:
                extra = text_file.readline(256 * 1024)
                if extra == "" or extra.endswith("\n") or extra.endswith("\r"):
                    break
            stats.skipped_long += 1
            return SKIPPED
        parts.append(chunk)
        if chunk.endswith("\n") or chunk.endswith("\r"):
            break
        if size > max_chars:
            while True:
                extra = text_file.readline(256 * 1024)
                if extra == "" or extra.endswith("\n") or extra.endswith("\r"):
                    break
            stats.skipped_long += 1
            return SKIPPED
    return "".join(parts).rstrip("\r\n")


def iter_lines(text_file: io.TextIOBase, max_chars: int, stats: Stats) -> Iterator[object]:
    while True:
        line = read_logical_line(text_file, max_chars, stats)
        if line is None:
            return
        yield line


_FMT_LINE = re.compile(
    r'^(\d+)\s+(\S+)\s+(\d+)\s+(\d+)\s+("(?:\\.|[^"])*")\s+(\d+)\s+(\S+)\s+("(?:\\.|[^"])*"|\S+)\s*$'
)


@dataclass
class FmtColumn:
    name: str
    prefix: int
    length: int
    terminator: str


def _unquote_fmt(token: str) -> str:
    body = token[1:-1] if len(token) >= 2 and token[0] == token[-1] == '"' else token
    try:
        return body.encode("utf-8").decode("unicode_escape")
    except UnicodeDecodeError as exc:
        raise ExportError(f"Не читается завершитель в форматном файле: {token}") from exc


def parse_fmt(path: str) -> tuple[list[FmtColumn], list[str]]:
    if not path:
        raise ExportError("Выберите форматный файл .fmt.")
    if not os.path.isfile(path):
        raise ExportError("Форматный файл не найден.")
    try:
        content = Path(path).read_text(encoding="utf-8-sig", errors="replace")
    except OSError as exc:
        raise ExportError(f"Не удалось прочитать форматный файл: {exc}") from exc
    lines = [line.strip() for line in content.splitlines() if line.strip()]
    if len(lines) < 3:
        raise ExportError("Форматный файл слишком короткий.")
    try:
        declared = int(lines[1])
    except ValueError as exc:
        raise ExportError("Во второй строке форматного файла должно быть число колонок.") from exc
    columns: list[FmtColumn] = []
    notes: list[str] = []
    for line in lines[2 : 2 + declared]:
        match = _FMT_LINE.match(line)
        if match is None:
            raise ExportError(f"Не разобрана строка форматного файла: {line}")
        columns.append(
            FmtColumn(
                name=match.group(7),
                prefix=int(match.group(3)),
                length=int(match.group(4)),
                terminator=_unquote_fmt(match.group(5)),
            )
        )
    if len(columns) != declared:
        notes.append("В форматном файле меньше строк, чем заявлено во второй строке.")
    if any(column.prefix != 0 for column in columns):
        raise ExportError(
            "В форматном файле есть префикс длины. Это native-файл bcp, не текст."
        )
    return columns, notes


def layout_from_fmt(path: str, encoding: str) -> Layout:
    columns, notes = parse_fmt(path)
    if not columns:
        raise ExportError("В форматном файле нет колонок.")
    names = [column.name for column in columns]
    last = columns[-1].terminator
    field_terms = [column.terminator for column in columns[:-1]]
    newline_terms = {"", "\n", "\r", "\r\n"}
    fixed = all(term == "" for term in field_terms) and all(column.length > 0 for column in columns)
    if fixed and last not in newline_terms and last != "":
        notes.append("Конец записи в форматном файле не перевод строки. Он снят с уже прочитанной записи.")
    if fixed:
        notes.append("Формат: фиксированная ширина полей из .fmt.")
        return Layout(
            encoding=encoding,
            mode="fixed",
            column_names=names,
            widths=[column.length for column in columns],
            notes=notes,
        )
    delimiter = field_terms[0] if field_terms else "\t"
    if field_terms and any(term != delimiter for term in field_terms):
        notes.append("В .fmt у колонок разные разделители. Поля режутся по порядку завершителей.")
        delimiter = "\t"
    notes.append(f"Формат: разделитель {describe_delimiter(delimiter)} из .fmt. Имена колонок взяты оттуда.")
    return Layout(
        encoding=encoding,
        mode="delimited",
        delimiter=delimiter,
        header_from_file=False,
        column_names=names,
        notes=notes,
    )


def is_noise(line: str) -> bool:
    text = line.strip()
    if not text:
        return True
    low = text.lower()
    if low.startswith("changed database context"):
        return True
    if low.startswith("контекст базы данных"):
        return True
    if low.startswith("network packet size"):
        return True
    if low.startswith("clock time"):
        return True
    if text.startswith("(") and text.endswith(")"):
        if any(token in low for token in ("row", "строк", "affected", "обработ", "затронут")):
            return True
    return False


def dash_spans(line: str) -> list[tuple[int, int | None]] | None:
    """Границы колонок отчёта sqlcmd/SSMS.

    Пробел между группами дефисов относится к левой колонке: пустое поле
    остаётся пустым и не сдвигает соседей влево. Последняя колонка идёт
    до конца строки, чтобы длинное значение не обрезалось и не слипалось
    со следующей. Колонка шириной в один символ тоже учитывается.
    """
    if not line or any(char not in " \t-" for char in line):
        return None
    runs: list[tuple[int, int]] = []
    index = 0
    length = len(line)
    while index < length:
        if line[index] == "-":
            end = index + 1
            while end < length and line[end] == "-":
                end += 1
            runs.append((index, end))
            index = end
        else:
            index += 1
    if not runs:
        return None
    longest = max(end - start for start, end in runs)
    if longest < 2 and len(runs) < 2:
        return None
    spans: list[tuple[int, int | None]] = []
    for position, (start, _end) in enumerate(runs):
        if position + 1 < len(runs):
            spans.append((start, runs[position + 1][0]))
        else:
            spans.append((start, None))
    return spans


def slice_spans(line: str, spans: list[tuple[int, int | None]]) -> list[str]:
    fields: list[str] = []
    for start, end in spans:
        if start >= len(line):
            fields.append("")
            continue
        piece = line[start:] if end is None else line[start:end]
        fields.append(piece.strip(" \t"))
    return fields


_WRAP_WIDTHS = {80, 100, 120, 132, 160, 200, 240, 256, 512, 1024, 2048, 4096, 8192, 16384, 32767, 65535}


def looks_like_quoted_csv(lines: list[str], delimiter: str) -> tuple[bool, bool]:
    """Кавычки считаются квалификатором, только если они обрамляют поля.

    Случайная кавычка внутри значения не должна склеивать хвост строки
    и следующую запись.
    """
    if not delimiter:
        return False, False
    pattern = re.compile(
        rf'(?:^|{re.escape(delimiter)})"(?:[^"]|"")*"(?={re.escape(delimiter)}|$)'
    )
    framed = 0
    bare = 0
    odd_lines = 0
    for line in lines[:40]:
        if line.count('"') % 2 == 1:
            odd_lines += 1
        found = len(pattern.findall(line))
        if found:
            framed += found
        elif '"' in line:
            bare += 1
    use_quotes = framed >= 1 and framed >= bare
    return use_quotes, use_quotes and odd_lines >= 1


def parse_fields(text: str, delimiter: str, qualifier: str) -> tuple[list[str], bool]:
    if not qualifier:
        return text.split(delimiter), False
    fields: list[str] = []
    buf: list[str] = []
    in_quote = False
    index = 0
    length = len(text)
    quote_len = len(qualifier)
    delim_len = len(delimiter)
    while index < length:
        if in_quote:
            if text.startswith(qualifier, index):
                if text.startswith(qualifier + qualifier, index):
                    buf.append(qualifier)
                    index += quote_len * 2
                    continue
                in_quote = False
                index += quote_len
                continue
            buf.append(text[index])
            index += 1
            continue
        if text.startswith(qualifier, index) and not buf:
            in_quote = True
            index += quote_len
            continue
        if text.startswith(delimiter, index):
            fields.append("".join(buf))
            buf = []
            index += delim_len
            continue
        buf.append(text[index])
        index += 1
    fields.append("".join(buf))
    return fields, in_quote


def fit_width(fields: list[str], width: int, delimiter: str, stats: Stats) -> list[str]:
    if width <= 0:
        width = max(1, len(fields))
    if len(fields) == width:
        return fields
    stats.ragged += 1
    if len(fields) > width:
        if width == 1:
            return [delimiter.join(fields)]
        head = fields[: width - 1]
        head.append(delimiter.join(fields[width - 1 :]))
        return head
    return fields + [""] * (width - len(fields))


def split_fixed(line: str, widths: list[int], encoding: str) -> list[str]:
    codec = _byte_codec(encoding)
    try:
        raw = line.encode(codec)
    except UnicodeEncodeError:
        raw = line.encode(codec, errors="replace")
    fields: list[str] = []
    pos = 0
    for width in widths:
        piece = raw[pos : pos + width]
        pos += width
        fields.append(piece.decode(codec, errors="replace").strip())
    return fields


def _byte_codec(encoding: str) -> str:
    normalized = encoding.lower().replace("_", "-")
    if normalized in {"utf-16", "utf-16-le"}:
        return "utf-16-le"
    if normalized == "utf-8-sig":
        return "utf-8"
    return encoding


def _numeric_token(value: str) -> bool:
    if not value:
        return False
    body = value[1:] if value[0] == "-" else value
    if body.isdigit():
        return True
    if body.count(".") == 1:
        left, right = body.split(".")
        return bool(left) and left.isdigit() and right.isdigit()
    return False


def _looks_like_header(first: list[str], second: list[str]) -> bool:
    if not first or not second or len(first) != len(second):
        return False
    if any(not item.strip() for item in first):
        return False
    first_numbers = sum(1 for item in first if _numeric_token(item))
    second_numbers = sum(1 for item in second if _numeric_token(item))
    return first_numbers == 0 and second_numbers > 0


def _choose_delimiter(lines: list[str]) -> tuple[str, bool]:
    usable = [line for line in lines if line.strip() and dash_spans(line) is None and not is_noise(line)]
    usable = usable[:40]
    if not usable:
        return "\t", False
    best = "\t"
    best_score = -1
    sure = False
    for delimiter in ("\t", "|", ";", ","):
        counts = [line.count(delimiter) for line in usable]
        if min(counts) < 1:
            continue
        stable = max(counts) == min(counts)
        score = (20 if stable else 0) + counts[0]
        if delimiter == "\t":
            score += 5
        if score > best_score:
            best = delimiter
            best_score = score
            sure = stable
    return best, sure and best_score >= 0


def _content_lines(text: str) -> list[str]:
    return [line for line in text.splitlines() if line.strip()]


def decide(sample: str, settings: Settings, encoding: str) -> Layout:
    notes = [f"Кодировка: {encoding}"]
    lines = sample.splitlines()
    if settings.mode in {"auto", "report"}:
        for index, line in enumerate(lines[:80]):
            spans = dash_spans(line)
            if spans and index > 0 and not is_noise(lines[index - 1]):
                notes.append(
                    "Режим: отчёт sqlcmd / SSMS. Ширина колонок взята из строки дефисов."
                )
                synthetic = settings.header == "no"
                if synthetic:
                    notes.append("Имена колонок заменены на «Колонка N» по вашему выбору.")
                return Layout(
                    encoding=encoding,
                    mode="report",
                    synthetic_header=synthetic,
                    header_from_file=not synthetic,
                    notes=notes,
                )
        if settings.mode == "report":
            raise ExportError(
                "В начале файла нет строки из дефисов. Это не отчёт sqlcmd или SSMS, либо образец её не захватил."
            )

    delimiter = settings.delimiter
    sure = True
    if delimiter == "auto":
        delimiter, sure = _choose_delimiter(lines)
        if not sure:
            notes.append("Разделитель угадан неуверенно. Если колонки съехали, выберите его вручную.")
    qualifier = settings.qualifier
    multiline = False
    if settings.mode == "csv" and not qualifier:
        qualifier = '"'
        multiline = True
    if settings.mode == "auto" and not qualifier and delimiter == ",":
        content_lines = [line for line in lines if line.strip() and dash_spans(line) is None][:40]
        use_quotes, _odd_quotes = looks_like_quoted_csv(content_lines, delimiter)
        if use_quotes:
            qualifier = '"'
    notes.append(f"Режим: текст с разделителем ({describe_delimiter(delimiter)}).")
    if qualifier:
        notes.append(
            "Кавычки обрамляют поля."
            if not multiline
            else "Кавычки обрамляют поля, в том числе с переносом строки."
        )

    content = [line for line in lines if line.strip()][:2]
    header_from_file = settings.header == "yes"
    if settings.header == "auto" and len(content) >= 2 and not qualifier:
        first, _closed_first = parse_fields(content[0], delimiter, "")
        second, _closed_second = parse_fields(content[1], delimiter, "")
        header_from_file = _looks_like_header(first, second)
        if header_from_file:
            notes.append("Первая строка принята за имена колонок. Если это данные, выберите «Без имён».")
        else:
            notes.append("Имён колонок в файле не видно. Заголовки будут «Колонка N».")
    elif settings.header == "no":
        notes.append("Первая строка пойдёт в данные, без отдельной строки имён.")
    return Layout(
        encoding=encoding,
        mode="delimited",
        delimiter=delimiter,
        qualifier=qualifier,
        header_from_file=header_from_file,
        allow_multiline_quotes=multiline,
        notes=notes,
    )


def build_layout(path: str, settings: Settings) -> Layout:
    if not os.path.isfile(path):
        raise ExportError("Исходный файл не найден.")
    encoding = resolve_encoding(path, settings.encoding)
    if settings.mode == "fmt":
        layout = layout_from_fmt(settings.fmt_path, encoding)
        layout.notes.insert(0, f"Кодировка: {encoding}")
        return layout
    sample = _read_sample(path, encoding)
    return decide(sample, settings, encoding)


def _read_sample(path: str, encoding: str) -> str:
    source = Source(path, encoding)
    try:
        text = source.text.read(SAMPLE_BYTES)
        extra = source.text.read(1)
        if extra:
            newline = text.rfind("\n")
            if newline >= 0:
                text = text[:newline]
        return text
    finally:
        source.close()


_DELETE_INVISIBLE = str.maketrans("", "", "\u200b\u200c\u200d\ufeff\u2060\u180e\u00ad")
_SPACE_LIKE = str.maketrans(
    {
        "\u00a0": " ",
        "\u2000": " ",
        "\u2001": " ",
        "\u2002": " ",
        "\u2003": " ",
        "\u2004": " ",
        "\u2005": " ",
        "\u2006": " ",
        "\u2007": " ",
        "\u2008": " ",
        "\u2009": " ",
        "\u200a": " ",
        "\u202f": " ",
        "\u205f": " ",
        "\u3000": " ",
    }
)
_SPACE_LIKE_KEEP_NBSP = str.maketrans(
    {key: " " for key in "\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2007\u2008\u2009\u200a\u202f\u205f\u3000"}
)
_INVISIBLE_NAMES = {
    "\u00a0": "неразрывный пробел",
    "\u2007": "цифровой пробел",
    "\u202f": "узкий пробел",
    "\u200b": "пробел нулевой ширины",
    "\u200c": "невидимый соединитель",
    "\u200d": "невидимый соединитель",
    "\ufeff": "метка BOM",
    "\u2060": "неразрывный знак нулевой ширины",
    "\u00ad": "мягкий перенос",
}
_SCAN_DELIMITERS = ("\t", ",", ";", "|", "\u00a0", "  ")


def normalize_text(text: str, *, keep_nbsp: bool = False) -> str:
    """Убирает невидимые символы. Неразрывный пробел остаётся, если это разделитель."""
    if not text:
        return text
    text = text.translate(_DELETE_INVISIBLE)
    return text.translate(_SPACE_LIKE_KEEP_NBSP if keep_nbsp else _SPACE_LIKE)


def _normalize_names(names: list[str], synthetic: bool) -> list[str]:
    if synthetic or not names:
        width = max(1, len(names))
        return [f"Колонка {index + 1}" for index in range(width)]
    cleaned: list[str] = []
    for index, name in enumerate(names):
        text = name.replace("\ufeff", "").strip()
        cleaned.append(text or f"Колонка {index + 1}")
    return cleaned


def _strip_mark(fields: list[str]) -> list[str]:
    if fields and fields[0].startswith("\ufeff"):
        fields[0] = fields[0].lstrip("\ufeff")
    return fields


class RowStream:
    def __init__(self, lines: Iterator[object], layout: Layout, settings: Settings, stats: Stats) -> None:
        self.lines = lines
        self.layout = layout
        self.settings = settings
        self.stats = stats
        self.header: list[str] = []
        self.restart = False
        self._pending: list[str] | None = None
        self._spans: list[tuple[int, int | None]] | None = None
        self._row_width = 0
        self._wrap_chunks: tuple[str, ...] = ()
        self._pushback: object | None = None
        self._report_started = False

    def _pull(self) -> object | None:
        if self._pushback is not None:
            item = self._pushback
            self._pushback = None
            return item
        item = next(self.lines, None)
        if isinstance(item, str):
            return normalize_text(item, keep_nbsp=self.layout.delimiter == "\u00a0")
        return item

    def _clean_fields(self, fields: list[str]) -> list[str]:
        # Пробел и невидимый пробел по краю ячейки — не значение. Иначе текст в Excel уезжает вправо.
        return [field.strip(" \t") for field in fields]

    def _drop_indent_column(self, text: str, fields: list[str]) -> list[str]:
        delimiter = self.layout.delimiter
        if not delimiter or len(fields) < 2 or fields[0] != "":
            return fields
        if text.startswith(delimiter):
            return fields
        return fields[1:]

    def _split_record(self, text: str) -> tuple[list[str], bool]:
        delimiter = self.layout.delimiter
        qualifier = self.layout.qualifier
        if delimiter == "  ":
            body = text.strip(" ")
            fields = [part.strip(" ") for part in re.split(r" {2,}", body)] if body else []
            return self._clean_fields(fields), False
        if not qualifier:
            fields = self._drop_indent_column(text, self._clean_fields(text.split(delimiter)))
            return fields, False
        fields, unclosed = parse_fields(text, delimiter, qualifier)
        if unclosed and not self.layout.allow_multiline_quotes:
            fields = self._drop_indent_column(text, self._clean_fields(text.split(delimiter)))
            return fields, False
        return self._drop_indent_column(text, self._clean_fields(fields)), unclosed

    def _join_dash_line(self, first: str) -> str:
        parts = [first]
        while True:
            nxt = self._pull()
            if nxt is None or nxt is SKIPPED:
                break
            text = str(nxt)
            if dash_spans(text) is None:
                self._pushback = nxt
                break
            parts.append(text)
        self._wrap_chunks = tuple(parts)
        joined = "".join(parts)
        self._row_width = len(joined)
        spans = dash_spans(joined)
        return joined if spans else first

    def _unwrap_data_line(self, text: str) -> str:
        if len(self._wrap_chunks) < 2 or self._row_width <= 0 or len(text) >= self._row_width:
            return text
        if len(text) not in _WRAP_WIDTHS or text.endswith(" "):
            return text
        parts = [text]
        while len("".join(parts)) < self._row_width:
            nxt = self._pull()
            if nxt is None or nxt is SKIPPED:
                break
            nxt_text = str(nxt)
            if (
                is_noise(nxt_text)
                or dash_spans(nxt_text) is not None
                or len(nxt_text) >= self._row_width
            ):
                self._pushback = nxt
                break
            parts.append(nxt_text)
            if len(nxt_text) not in _WRAP_WIDTHS:
                break
        return "".join(parts)

    def open(self) -> None:
        if self.layout.mode == "report":
            self._prime_report()
        elif self.layout.mode == "fixed":
            self.header = _normalize_names(self.layout.column_names, False)
        else:
            self._prime_delimited()
        self._lock_columns()
        if len(self.header) > 16384:
            raise ExportError("В строке больше 16 384 колонок. Столько лист Excel не принимает.")

    def rows(self) -> Iterator[list[str]]:
        if self.layout.mode == "report":
            yield from self._rows_report()
        elif self.layout.mode == "fixed":
            yield from self._rows_fixed()
        else:
            yield from self._rows_delimited()

    def _prime_delimited(self) -> None:
        while True:
            line = self._pull()
            if line is None:
                self.header = _normalize_names(self.layout.column_names, not self.layout.column_names)
                return
            if line is SKIPPED or line == "":
                continue
            fields, unclosed = self._split_record(str(line))
            if unclosed:
                raise ExportError("В образце не закрыта кавычка. Проверьте квалификатор текста.")
            fields = _strip_mark(fields)
            if self.layout.column_names and not self.layout.header_from_file:
                self.header = _normalize_names(self.layout.column_names, False)
                self._pending = fit_width(fields, len(self.header), self.layout.delimiter, self.stats)
                return
            if self.layout.header_from_file and (
                not self.layout.locked_columns or len(fields) == self.layout.locked_columns
            ):
                self.header = _normalize_names(fields, False)
                self._lock_columns()
                return
            width = self.layout.locked_columns or max(1, len(fields))
            self.header = [f"Колонка {index + 1}" for index in range(width)]
            self._pending = fit_width(fields, len(self.header), self.layout.delimiter, self.stats)
            return

    def _lock_columns(self) -> None:
        locked = self.layout.locked_columns
        if locked <= 0 or not self.header:
            return
        if len(self.header) < locked:
            start = len(self.header)
            self.header.extend(f"Колонка {start + index + 1}" for index in range(locked - start))

    def _rows_delimited(self) -> Iterator[list[str]]:
        width = len(self.header)
        if self._pending is not None:
            pending = self._pending
            self._pending = None
            yield pending
        carry: str | None = None
        delimiter = self.layout.delimiter
        while True:
            line = self._pull()
            if line is None:
                break
            if line is SKIPPED:
                continue
            if line == "" and carry is None:
                continue
            text = str(line) if carry is None else carry + "\n" + str(line)
            if len(text) > self.settings.max_line_chars:
                self.stats.skipped_long += 1
                carry = None
                continue
            fields, unclosed = self._split_record(text)
            if unclosed:
                carry = text
                continue
            carry = None
            yield fit_width(fields, width, delimiter, self.stats)
        if carry:
            fields = self._clean_fields(carry.split(delimiter))
            self.stats.ragged += 1
            yield fit_width(fields, width, delimiter, self.stats)

    def _prime_report(self) -> None:
        previous = ""
        while True:
            line = self._pull()
            if line is None:
                break
            if line is SKIPPED:
                previous = ""
                continue
            text = str(line)
            if dash_spans(text) is not None and previous.strip() and not is_noise(previous):
                spans = dash_spans(self._join_dash_line(text))
                if not spans:
                    previous = ""
                    continue
                self._spans = spans
                self._report_started = True
                self.header = _normalize_names(slice_spans(previous, spans), self.layout.synthetic_header)
                return
            if is_noise(text):
                previous = ""
                continue
            previous = text
        raise ExportError("В файле не найден заголовок отчёта со строкой из дефисов.")

    def _rows_report(self) -> Iterator[list[str]]:
        # Строка держится до следующей: повтор заголовка sqlcmd не должен
        # попасть в данные, а в constant_memory уже записанную строку не стереть.
        held: str | None = None
        while True:
            line = self._pull()
            if line is None:
                break
            if line is SKIPPED:
                continue
            text = str(line)
            if dash_spans(text) is not None and held is not None and held.strip() and not is_noise(held):
                spans = dash_spans(self._join_dash_line(text))
                if spans:
                    names = _normalize_names(slice_spans(held, spans), self.layout.synthetic_header)
                    self._spans = spans
                    if names != self.header:
                        self.header = names
                        self.restart = True
                held = None
                continue
            if held is not None and self._spans is not None and not is_noise(held):
                yield slice_spans(self._unwrap_data_line(held), self._spans)
            held = None
            if is_noise(text):
                if text.strip():
                    self._spans = None
                continue
            held = text
        if held is not None and self._spans is not None and not is_noise(held):
            yield slice_spans(self._unwrap_data_line(held), self._spans)

    def _rows_fixed(self) -> Iterator[list[str]]:
        widths = self.layout.widths
        encoding = self.layout.encoding
        expected = len(widths)
        for line in self.lines:
            if line is SKIPPED or line == "":
                continue
            fields = split_fixed(str(line), widths, encoding)
            if len(fields) != expected:
                self.stats.ragged += 1
                fields = fit_width(fields, expected, "", self.stats)
            yield fields


def form_value(text: str, stats: Stats):
    if text is None:
        return None
    if text == "":
        return None
    if len(text) > 32767:
        stats.truncated += 1
        text = text[:32767]
    sign = text[0] == "-"
    body = text[1:] if sign else text
    if body.isdigit():
        if len(body) > 1 and body[0] == "0":
            return text
        if len(body) > 15:
            return text
        number = int(body)
        return -number if sign else number
    if body.count(".") == 1:
        left, right = body.split(".")
        if left.isdigit() and right.isdigit() and left and not (len(left) > 1 and left[0] == "0"):
            if len(left) + len(right) <= 15:
                return float(text)
    if len(text) >= 10 and text[4] == "-" and text[7] == "-":
        parsed = _form_date(text)
        if parsed is not None:
            return parsed
    return text


def _form_date(text: str):
    try:
        if len(text) == 10:
            return dt.date.fromisoformat(text)
        if len(text) >= 19 and text[10] in " T" and text[13] == ":" and text[16] == ":":
            value = dt.datetime.strptime(text[:19].replace("T", " "), "%Y-%m-%d %H:%M:%S")
            if len(text) > 19 and text[19] == ".":
                digits: list[str] = []
                for char in text[20:]:
                    if char.isdigit():
                        digits.append(char)
                    else:
                        break
                if digits:
                    micro = int("".join(digits)[:6].ljust(6, "0"))
                    value = value.replace(microsecond=micro)
            return value
    except ValueError:
        return None
    return None


class BookWriter:
    def __init__(self, output: str, rows_per_file: int) -> None:
        self.base = Path(output)
        self.limit = min(max(1, rows_per_file), EXCEL_MAX_DATA_ROWS)
        self.files: list[str] = []
        self.workbook = None
        self.sheet = None
        self.header_format = None
        self.date_format = None
        self.time_format = None
        self.row_index = 0
        self.rows_in_file = 0
        self.file_index = 0
        self.header: list[str] = []

    def start(self, header: list[str]) -> None:
        self.header = header
        self._open_next()

    def roll(self, header: list[str]) -> None:
        self.header = header
        self._open_next()

    def write(self, values: list[object], stats: Stats) -> None:
        if self.sheet is None or self.rows_in_file >= self.limit:
            self._open_next()
        wrote = False
        for column, value in enumerate(values):
            if value is None or value == "":
                continue
            if isinstance(value, bool):
                self.sheet.write_number(self.row_index, column, int(value))
            elif isinstance(value, int) and not isinstance(value, bool):
                self.sheet.write_number(self.row_index, column, value)
            elif isinstance(value, float):
                self.sheet.write_number(self.row_index, column, value)
            elif isinstance(value, dt.datetime):
                self.sheet.write_datetime(self.row_index, column, value, self.date_format)
            elif isinstance(value, dt.date):
                self.sheet.write_datetime(self.row_index, column, value, self.date_format)
            elif isinstance(value, dt.time):
                self.sheet.write_datetime(self.row_index, column, value, self.time_format)
            else:
                text = str(value)
                if len(text) > 32767:
                    stats.truncated += 1
                    text = text[:32767]
                self.sheet.write_string(self.row_index, column, text)
            wrote = True
        if not wrote:
            self.sheet.write_blank(self.row_index, 0, None)
        self.row_index += 1
        self.rows_in_file += 1

    def close(self) -> None:
        if self.workbook is not None:
            self.workbook.close()
            self.workbook = None
            self.sheet = None

    def _open_next(self) -> None:
        self.close()
        self.file_index += 1
        path = self.base if self.file_index == 1 else self.base.with_name(
            f"{self.base.stem}_{self.file_index:03d}{self.base.suffix}"
        )
        workbook = xlsxwriter.Workbook(
            str(path),
            {
                "constant_memory": True,
                "strings_to_formulas": False,
                "strings_to_urls": False,
                "strings_to_numbers": False,
                "default_date_format": "yyyy-mm-dd hh:mm:ss",
            },
        )
        workbook.use_zip64()
        sheet = workbook.add_worksheet("Данные")
        header_format = workbook.add_format(
            {
                "bold": True,
                "bg_color": "#F4E7D4",
                "bottom": 1,
                "font_name": "Calibri",
            }
        )
        for column, name in enumerate(self.header):
            sheet.write_string(0, column, name, header_format)
            width = min(42, max(10, len(name) + 2))
            sheet.set_column(column, column, width)
        if self.header:
            sheet.freeze_panes(1, 0)
        self.workbook = workbook
        self.sheet = sheet
        self.header_format = header_format
        self.date_format = workbook.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"})
        self.time_format = workbook.add_format({"num_format": "hh:mm:ss"})
        self.row_index = 1
        self.rows_in_file = 0
        self.files.append(str(path))


TEXT_SUFFIXES = {".txt", ".rpt", ".csv", ".dat", ".bcp"}


def plan_folder(folder: str) -> list[tuple[str, str]]:
    """Все текстовые выгрузки в папке и путь книги рядом с каждой."""
    root = Path(folder)
    if not root.is_dir():
        raise ExportError("Папка не найдена.")
    files = sorted(
        (
            path
            for path in root.iterdir()
            if path.is_file() and path.suffix.lower() in TEXT_SUFFIXES and not path.name.startswith("~$")
        ),
        key=lambda path: path.name.lower(),
    )
    if not files:
        raise ExportError("В папке нет файлов txt, rpt, csv, dat или bcp.")
    same_stem: dict[str, int] = {}
    for path in files:
        key = path.stem.lower()
        same_stem[key] = same_stem.get(key, 0) + 1
    pairs: list[tuple[str, str]] = []
    for path in files:
        if same_stem[path.stem.lower()] == 1:
            target = path.with_suffix(".xlsx")
        else:
            target = path.parent / f"{path.name}.xlsx"
        pairs.append((str(path), str(target)))
    return pairs


def _quoted_mark(text: str, delimiter: str) -> int:
    """Сколько разделителей снаружи кавычек. Запятая внутри поля не плодит колонку."""
    if '"' not in text:
        return text.count(delimiter)
    fields, unclosed = parse_fields(text, delimiter, '"')
    if unclosed:
        return text.count(delimiter)
    return max(0, len(fields) - 1)


def _modal_split(counts: dict[int, int]) -> tuple[int, int, int]:
    if not counts:
        return 0, 0, 0
    total = sum(counts.values())
    positive = {count: lines for count, lines in counts.items() if count > 0}
    if not positive:
        return 0, 0, total
    count, lines = max(positive.items(), key=lambda item: (item[1], item[0]))
    return count, lines, total


def _best_scanned_delimiter(hist: dict[str, dict[int, int]]) -> tuple[str, bool, int]:
    best = "\t"
    best_score = -1.0
    best_fields = 0
    sure = False
    for delimiter, counts in hist.items():
        mode_count, mode_lines, total = _modal_split(counts)
        if mode_count < 1 or total <= 0:
            continue
        stability = mode_lines / total
        if stability < 0.55:
            continue
        score = stability * 1000 + mode_count
        if delimiter == "\t":
            score += 5
        if score > best_score:
            best = delimiter
            best_score = score
            best_fields = mode_count + 1
            sure = stability >= 0.8
    return best, sure, best_fields


def analyze_file(
    path: str,
    settings: Settings,
    progress: Callable[[Progress], None] | None = None,
    cancelled: Callable[[], bool] | None = None,
) -> Layout:
    """Первый проход: весь файл, в памяти только счётчики и несколько строк."""
    if not os.path.isfile(path):
        raise ExportError("Исходный файл не найден.")
    encoding = resolve_encoding(path, settings.encoding)
    if settings.mode == "fmt":
        layout = layout_from_fmt(settings.fmt_path, encoding)
        layout.notes.insert(0, f"Кодировка: {encoding}")
        return layout

    source = Source(path, encoding)
    stats = Stats()
    histogram: dict[str, dict[int, int]] = {delimiter: {} for delimiter in _SCAN_DELIMITERS}
    invisible: dict[str, int] = {}
    sample_lines: list[str] = []
    previous = ""
    dash_line = ""
    seen = 0
    try:
        progress and progress(
            Progress(stage="probe", fraction=0.01, message="Анализ файла", file_size=source.size)
        )
        for line in iter_lines(source.text, settings.max_line_chars, stats):
            if line is SKIPPED:
                continue
            seen += 1
            if cancelled and seen % PROGRESS_EVERY == 0 and cancelled():
                raise Cancelled("Остановлено на анализе.")
            raw = str(line)
            for char, name in _INVISIBLE_NAMES.items():
                amount = raw.count(char)
                if amount:
                    invisible[name] = invisible.get(name, 0) + amount
            text = normalize_text(raw, keep_nbsp=True)
            if seen % PROGRESS_EVERY == 0 and progress:
                progress(
                    Progress(
                        stage="probe",
                        fraction=_fraction(source.counter.read_bytes, source.size),
                        message="Анализ разделителей и невидимых символов",
                        rows=seen,
                        bytes_read=source.counter.read_bytes,
                        file_size=source.size,
                    )
                )
            if is_noise(text):
                previous = ""
                continue
            if not dash_line and seen <= 200 and dash_spans(text) is not None and previous.strip():
                dash_line = text[:8000]
            if dash_spans(text) is None:
                if len(sample_lines) < 12 and len(text) <= 4000:
                    sample_lines.append(text)
                for delimiter, bucket in histogram.items():
                    if delimiter == "  ":
                        count = len(re.findall(r" {2,}", text))
                    elif delimiter in {",", ";"}:
                        count = _quoted_mark(text, delimiter)
                    else:
                        count = text.count(delimiter)
                    bucket[count] = bucket.get(count, 0) + 1
            previous = text[:4000]
    finally:
        source.close()

    notes = [f"Кодировка: {encoding}", f"Анализ всего файла: {_format_scan_rows(seen)} строк."]
    for name, amount in invisible.items():
        notes.append(f"Найден {name}: {amount}.")
    if settings.mode in {"auto", "report"} and dash_line:
        notes.append("Режим: отчёт sqlcmd / SSMS. Ширина колонок снята со строки дефисов во всём файле.")
        synthetic = settings.header == "no"
        if synthetic:
            notes.append("Имена колонок заменены на «Колонка N».")
        return Layout(
            encoding=encoding,
            mode="report",
            synthetic_header=synthetic,
            header_from_file=not synthetic,
            notes=notes,
        )
    if settings.mode == "report":
        raise ExportError("В файле нет строки из дефисов. Это не отчёт sqlcmd или SSMS.")

    guessed, sure, fields = _best_scanned_delimiter(histogram)
    delimiter = settings.delimiter if settings.delimiter not in {"", "auto"} else guessed
    if settings.delimiter in {"", "auto"} and not sure:
        notes.append("Разделитель по всему файлу неоднородный. Если колонки съедут, выберите его вручную.")
    locked = 0
    if delimiter in histogram:
        mode_count, _lines, _total = _modal_split(histogram[delimiter])
        if mode_count > 0:
            locked = mode_count + 1
    elif fields:
        locked = fields

    qualifier = settings.qualifier
    multiline = settings.mode == "csv"
    if settings.mode == "csv" and not qualifier:
        qualifier = '"'
    if settings.mode == "auto" and not qualifier and delimiter == ",":
        use_quotes, _odd = looks_like_quoted_csv(sample_lines, delimiter)
        if use_quotes:
            qualifier = '"'
    notes.append(f"Режим: текст с разделителем ({describe_delimiter(delimiter)}).")
    if locked:
        notes.append(f"Колонок по всему файлу: {locked}. Пустые поля между разделителями сохраняются.")
    notes.append("Пробел по краю ячейки снимается и не сдвигает строку вправо.")
    if qualifier:
        notes.append("Кавычки обрамляют поля." if not multiline else "Кавычки обрамляют поля, в том числе с переносом строки.")

    header_from_file = settings.header == "yes"
    if settings.header == "auto" and locked and sample_lines and not qualifier:
        matching = [line for line in sample_lines if line.count(delimiter) + 1 == locked]
        first_count = sample_lines[0].count(delimiter) + 1
        if first_count == locked and len(matching) >= 2:
            first = matching[0].split(delimiter)
            second = matching[1].split(delimiter)
            header_from_file = _looks_like_header(first, second)
            notes.append(
                "Первая подходящая строка принята за имена колонок."
                if header_from_file
                else "Имена колонок не найдены. Заголовки будут «Колонка N»."
            )
        else:
            notes.append("Первая строка не задаёт колонки. Имена будут «Колонка N».")
    elif settings.header == "no":
        notes.append("Первая строка пойдёт в данные, без отдельной строки имён.")
    return Layout(
        encoding=encoding,
        mode="delimited",
        delimiter=delimiter,
        qualifier=qualifier,
        header_from_file=header_from_file,
        allow_multiline_quotes=multiline,
        locked_columns=locked,
        notes=notes,
    )


def _format_scan_rows(value: int) -> str:
    return f"{value:,}".replace(",", " ")


def inspect(path: str, settings: Settings) -> Layout:
    layout = build_layout(path, settings)
    stats = Stats()
    source = Source(path, layout.encoding)
    try:
        stream = RowStream(iter_lines(source.text, settings.max_line_chars, stats), layout, settings, stats)
        stream.open()
        sample: list[list[str]] = []
        for row in stream.rows():
            sample.append(list(row))
            if len(sample) >= 8:
                break
        layout.preview_header = list(stream.header)
        layout.sample = sample
        return layout
    finally:
        source.close()


def _fraction(read_bytes: int, size: int) -> float:
    if size <= 0:
        return 0.0
    return min(read_bytes / size, 0.99)


def convert(
    path: str,
    output: str,
    settings: Settings,
    progress: Callable[[Progress], None],
    cancelled: Callable[[], bool],
    on_preview: Callable[[list[str], list[list[str]]], None] | None = None,
) -> ConvertResult:
    import time

    if not output.lower().endswith(".xlsx"):
        raise ExportError("Книга Excel должна заканчиваться на .xlsx.")
    if Path(path).resolve() == Path(output).resolve():
        raise ExportError("Файл назначения совпадает с исходным.")
    parent = Path(output).parent
    if not parent.exists():
        raise ExportError("Папка для книги Excel не существует.")
    if settings.qualifier and settings.delimiter not in {"", "auto"} and settings.qualifier == settings.delimiter:
        raise ExportError("Кавычка и разделитель не должны совпадать.")

    started = time.perf_counter()
    progress(Progress(stage="probe", fraction=0.02, message="Анализ всего файла"))
    layout = analyze_file(path, settings, progress, cancelled)
    if cancelled():
        raise Cancelled("Остановлено.")
    progress(
        Progress(
            stage="probe",
            fraction=1.0,
            message=" ; ".join(layout.notes[:3]) if layout.notes else "Формат определён",
        )
    )

    stats = Stats()
    source = Source(path, layout.encoding)
    writer: BookWriter | None = None
    rows_written = 0
    preview: list[list[str]] = []
    preview_sent = False
    try:
        stream = RowStream(
            iter_lines(source.text, settings.max_line_chars, stats),
            layout,
            settings,
            stats,
        )
        stream.open()
        if not stream.header:
            raise ExportError("В файле нет строк, которые можно разобрать.")
        writer = BookWriter(output, settings.rows_per_file)
        writer.start(stream.header)
        progress(
            Progress(
                stage="read",
                fraction=0.0,
                message="Чтение",
                file_size=source.size,
                files_done=1,
            )
        )
        for raw_row in stream.rows():
            if stream.restart:
                stream.restart = False
                if cancelled():
                    raise Cancelled("Остановлено. Уже собранная часть книги остаётся на диске.")
                progress(Progress(stage="write", fraction=_fraction(source.counter.read_bytes, source.size), message="Собираю текущую книгу", rows=rows_written, bytes_read=source.counter.read_bytes, file_size=source.size))
                writer.roll(stream.header)
            formed = [form_value(value, stats) for value in raw_row]
            writer.write(formed, stats)
            rows_written += 1
            if on_preview is not None and len(preview) < 8:
                preview.append(list(raw_row))
                if len(preview) == 8:
                    on_preview(list(stream.header), preview)
                    preview_sent = True
            if rows_written % PROGRESS_EVERY == 0:
                if cancelled():
                    raise Cancelled("Остановлено. Уже собранная часть книги остаётся на диске.")
                elapsed = max(time.perf_counter() - started, 1e-6)
                fraction = _fraction(source.counter.read_bytes, source.size)
                speed = rows_written / elapsed
                common = dict(
                    rows=rows_written,
                    bytes_read=source.counter.read_bytes,
                    file_size=source.size,
                    files_done=len(writer.files),
                    rows_per_sec=speed,
                    fraction=fraction,
                )
                progress(Progress(stage="read", message="Чтение файла", **common))
                progress(Progress(stage="parse", message="Разбор строк", **common))
                progress(Progress(stage="form", message="Формирование ячеек", **common))
                progress(Progress(stage="write", message="Запись книги", **common))
        if on_preview is not None and preview and not preview_sent:
            on_preview(list(stream.header), preview)
        if cancelled():
            raise Cancelled("Остановлено. Уже собранная часть книги остаётся на диске.")
        progress(
            Progress(
                stage="write",
                fraction=1,
                message="Упаковываю книгу Excel",
                rows=rows_written,
                bytes_read=source.size,
                file_size=source.size,
                files_done=len(writer.files),
            )
        )
        files = list(writer.files)
        columns = len(stream.header)
        writer.close()
        writer = None
    finally:
        if writer is not None:
            writer.close()
        source.close()

    elapsed = time.perf_counter() - started
    result = ConvertResult(
        rows=rows_written,
        files=files,
        elapsed=elapsed,
        truncated_cells=stats.truncated,
        skipped_long=stats.skipped_long,
        ragged_rows=stats.ragged,
        encoding=layout.encoding,
        mode=layout.mode,
        columns=columns,
    )
    progress(
        Progress(
            stage="done",
            fraction=1,
            message="Готово",
            rows=result.rows,
            bytes_read=source.size,
            file_size=source.size,
            files_done=len(result.files),
            rows_per_sec=result.rows / elapsed if elapsed else 0,
        )
    )
    return result


def run_self_check() -> None:
    import tempfile
    import time
    import tracemalloc

    assert parse_fields('a,"b,c",d', ",", '"') == (["a", "b,c", "d"], False)
    assert parse_fields('a,"b""c",d', ",", '"') == (["a", 'b"c', "d"], False)
    assert parse_fields("1\t\t3", "\t", "") == (["1", "", "3"], False)

    spans = dash_spans("--   ----")
    assert spans == [(0, 5), (5, None)]
    assert slice_spans("1    Ann", spans) == ["1", "Ann"]
    assert slice_spans("1       Kazan", dash_spans("-- ---- ----")) == ["1", "", "Kazan"]
    narrow = dash_spans("- ----")
    assert narrow is not None and len(narrow) == 2
    assert slice_spans("1 Ann", narrow) == ["1", "Ann"]
    assert form_value("00123", Stats()) == "00123"
    assert form_value("42", Stats()) == 42
    assert form_value("", Stats()) is None
    quote_text = 'a,b,c\n1,"x,y",3\n4,"5,6,7\n8,9,0\n'
    quote_layout = decide(quote_text, Settings(mode="auto", delimiter=",", header="yes"), "utf-8")
    assert quote_layout.qualifier == '"'
    assert quote_layout.allow_multiline_quotes is False
    quote_stream = RowStream(iter(quote_text.splitlines()), quote_layout, Settings(), Stats())
    quote_stream.open()
    quote_rows = list(quote_stream.rows())
    assert quote_rows[0] == ["1", "x,y", "3"], quote_rows
    assert quote_rows[1][0] == "4"
    assert quote_rows[2] == ["8", "9", "0"], quote_rows
    assert form_value("1998-08-12 00:00:00.000", Stats()) == dt.datetime(1998, 8, 12, 0, 0, 0)

    directory = tempfile.mkdtemp(prefix="alfa_excel_")
    try:
        scan_path = os.path.join(directory, "scan.txt")
        Path(scan_path).write_text(
            "h1\u00a0h2\u00a0h3\n"
            "1\u00a0\u00a02\n"
            "3\u00a0x\u00a04\n"
            "4\u00a0y\u00a05\n",
            encoding="utf-8",
        )
        scanned = analyze_file(scan_path, Settings(mode="auto", delimiter="auto", header="yes"))
        assert scanned.delimiter == "\u00a0", (scanned.delimiter, scanned.notes)
        assert scanned.locked_columns == 3, scanned.locked_columns
        assert any("неразрывный" in note for note in scanned.notes)
        held = Source(scan_path, "utf-8")
        try:
            scan_stream = RowStream(iter_lines(held.text, 100000, Stats()), scanned, Settings(), Stats())
            scan_stream.open()
            scan_rows = list(scan_stream.rows())
        finally:
            held.close()
        assert scan_rows[0] == ["1", "", "2"], scan_rows
        assert scan_rows[1] == ["3", "x", "4"], scan_rows
        pad_path = os.path.join(directory, "pad.txt")
        Path(pad_path).write_text("name\tage\n Ann\t10\n\t20\n1\t\n", encoding="utf-8")
        pad_layout = analyze_file(pad_path, Settings(mode="delimited", delimiter="\t", header="yes"))
        held_pad = Source(pad_path, "utf-8")
        try:
            pad_stream = RowStream(iter_lines(held_pad.text, 100000, Stats()), pad_layout, Settings(), Stats())
            pad_stream.open()
            pad_rows = list(pad_stream.rows())
        finally:
            held_pad.close()
        assert pad_rows[0] == ["Ann", "10"], pad_rows
        assert pad_rows[1] == ["", "20"], pad_rows
        assert pad_rows[2] == ["1", ""], pad_rows
        source = os.path.join(directory, "dump.txt")
        output = os.path.join(directory, "dump.xlsx")
        with open(source, "w", encoding="utf-8", newline="") as handle:
            handle.write("code\tname\tamount\twhen\n")
            for index in range(4000):
                handle.write(f"00{index}\tИмя {index}\t{index}.50\t1998-08-12 00:00:00.000\n")
        settings = Settings(encoding="utf-8", mode="delimited", delimiter="\t", header="yes", rows_per_file=1500)
        tracemalloc.start()
        result = convert(source, output, settings, lambda _event: None, lambda: False)
        _current, peak = tracemalloc.get_traced_memory()
        tracemalloc.stop()
        assert result.rows == 4000, result.rows
        assert len(result.files) == 3, result.files
        assert peak < 50 * 1024 * 1024, peak
        with zipfile.ZipFile(result.files[0]) as book:
            xml = book.read("xl/worksheets/sheet1.xml")
        assert b"inlineStr" in xml or "code".encode() in xml
        assert b">42<" in xml or b">0<" in xml or b">1.5<" in xml or b">1500<" in xml or b">0.5<" in xml

        report = os.path.join(directory, "report.rpt")
        report_out = os.path.join(directory, "report.xlsx")
        with open(report, "w", encoding="utf-8", newline="") as handle:
            handle.write("Changed database context to 'Demo'.\n")
            handle.write("ID   Name\n")
            handle.write("--   ----\n")
            handle.write("1    Ann\n")
            handle.write("7    Bo\n")
            handle.write("ID   Name\n")
            handle.write("--   ----\n")
            handle.write("8    Cy\n")
            handle.write("(3 rows affected)\n")
        report_result = convert(
            report,
            report_out,
            Settings(encoding="utf-8", mode="auto", header="auto"),
            lambda _event: None,
            lambda: False,
        )
        assert report_result.rows == 3, report_result.rows
        assert report_result.mode == "report"
        with zipfile.ZipFile(report_out) as book:
            report_xml = book.read("xl/worksheets/sheet1.xml")
        assert "Ann".encode() in report_xml

        quoted = os.path.join(directory, "quoted.csv")
        quoted_out = os.path.join(directory, "quoted.xlsx")
        with open(quoted, "w", encoding="utf-8", newline="") as handle:
            handle.write('name,note\n"Анна","a,b"\n')
        quoted_result = convert(
            quoted,
            quoted_out,
            Settings(encoding="utf-8", mode="csv", delimiter=",", qualifier='"', header="yes"),
            lambda _event: None,
            lambda: False,
        )
        assert quoted_result.rows == 1
        with zipfile.ZipFile(quoted_out) as book:
            quoted_xml = book.read("xl/worksheets/sheet1.xml")
        assert "a,b".encode() in quoted_xml

        fmt_path = os.path.join(directory, "table.fmt")
        fmt_data = os.path.join(directory, "table.txt")
        fmt_out = os.path.join(directory, "table.xlsx")
        fmt_path_text = "\n".join(
            [
                "14.0",
                "2",
                '1       SQLCHAR       0       0       "\\t"     1     Code     ""',
                '2       SQLCHAR       0       0       "\\r\\n"   2     Title    ""',
                "",
            ]
        )
        Path(fmt_path).write_text(fmt_path_text, encoding="utf-8")
        Path(fmt_data).write_text("10\tНасос\n", encoding="utf-8")
        fmt_result = convert(
            fmt_data,
            fmt_out,
            Settings(encoding="utf-8", mode="fmt", fmt_path=fmt_path),
            lambda _event: None,
            lambda: False,
        )
        assert fmt_result.rows == 1
        assert fmt_result.columns == 2

        wide = os.path.join(directory, "wide.txt")
        with open(wide, "wb") as handle:
            handle.write(b"\xff\xfe")
            handle.write("1\t2\r\n".encode("utf-16-le"))
        assert resolve_encoding(wide, "auto") == "utf-16"
        size = os.path.getsize(source)
        checker = Source(source, "utf-8")
        try:
            while checker.text.readline():
                pass
        finally:
            read_bytes = checker.counter.read_bytes
            checker.close()
        assert read_bytes == size, (read_bytes, size)
        folder = os.path.join(directory, "batch")
        os.makedirs(folder)
        Path(os.path.join(folder, "one.txt")).write_text("a\tb\n1\t2\n", encoding="utf-8")
        Path(os.path.join(folder, "two.csv")).write_text("a;b\n3;4\n", encoding="utf-8")
        Path(os.path.join(folder, "skip.xlsx")).write_bytes(b"")
        Path(os.path.join(folder, "note.md")).write_text("no", encoding="utf-8")
        planned = plan_folder(folder)
        assert [Path(item[0]).name for item in planned] == ["one.txt", "two.csv"]
        assert Path(planned[0][1]).name == "one.xlsx"
        assert Path(planned[1][1]).name == "two.xlsx"
        _ = time.perf_counter
    finally:
        import shutil

        shutil.rmtree(directory, ignore_errors=True)
    print("CONVERT_CHECK_OK", flush=True)
