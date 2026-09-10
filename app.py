from datetime import datetime, timezone
import base64
import html as html_lib
import io
import os

import pandas as pd
import streamlit as st
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from sqlalchemy import Column, Integer, String, create_engine, inspect, text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

try:
    import qrcode
    import qrcode.image.svg  # noqa: F401

    QR_AVAILABLE = True
except Exception:  # библиотека не установлена — QR просто не показываем
    QR_AVAILABLE = False


def utc_now_str():
    """Текущее время в UTC в едином формате записей."""
    return datetime.now(timezone.utc).strftime("%b %d, %Y %H:%M UTC")


# Русские заголовки колонок для отображения таблиц
COL_RU = {
    "id": "ID",
    "date": "Дата",
    "party": "Партия",
    "category": "Категория",
    "model": "Модель",
    "serial_number": "Серийный номер",
    "inv_number": "Инвентарный номер",
    "quantity": "Кол-во",
    "condition": "Состояние",
    "engineer": "Ответственный",
    "date_updated": "Обновлено",
    "equipment_info": "Техника",
    "from_where": "Откуда",
    "to_where": "Куда",
    "receiver": "Принял(а)",
    "action": "Действие",
    "details": "Что изменилось",
}


def validate_transfer_surnames(own_surname, receiver_surname, engineer_fio, match_fio=True):
    """Проверяет обязательные фамилии при перемещении. Возвращает (ok, текст ошибки).

    match_fio=True — фамилия исполнителя сверяется с ФИО из меню (режим инженера);
    match_fio=False — только обязательность полей (режим администратора).
    """
    own = (own_surname or "").strip()
    recv = (receiver_surname or "").strip()
    fio = (engineer_fio or "").strip()

    if not own:
        return False, "Укажите фамилию исполнителя — это поле обязательно."
    if not recv:
        return False, "Укажите фамилию принимающего — это поле обязательно."

    if not match_fio:
        return True, ""

    fio_words = [w.lower() for w in fio.split()]
    if not fio_words:
        return False, "Сначала укажите ваше ФИО в боковом меню слева."
    if own.lower() not in fio_words:
        return False, (
            f"Ваша фамилия «{own}» не совпадает с ФИО из бокового меню "
            f"(«{fio}»). Проверьте и повторите."
        )
    return True, ""


# --- ПРАВИЛА И РАСЧЁТЫ ---
# Лимит единиц техники одной категории в одной партии.
# Ноутбуки и сетевые фильтры — не более 2 шт, остальная оргтехника — не более 1 шт
# (моноблок, ПК, монитор, ИБП и т.д. попадают под общее правило «не более 1»).
LIMITS_SPECIAL = {"Ноутбук": 2, "Сетевой фильтр": 2}
DEFAULT_LIMIT = 1
BASE_PARTY = "База"
BAD_ROW_STYLE = "background-color: #ffe1e1; color: #a40000; font-weight: 600"


def category_limit(category):
    """Максимальное число единиц категории в одной партии."""
    return LIMITS_SPECIAL.get(category, DEFAULT_LIMIT)


def is_bad_condition(value):
    """True, если позиция помечена как неудовлетворительная."""
    return str(value).strip().lower() != "удовлетворительно"


def condition_text(value):
    """Читаемый текст состояния без служебного префикса."""
    text = str(value).strip()
    low = text.lower()
    for prefix in (
        "не удовлетворительно:",
        "неудовлетворительно:",
        "не удовлетворительно",
        "неудовлетворительно",
    ):
        if low.startswith(prefix):
            stripped = text[len(prefix):].strip(" :—-")
            return stripped or "неудовлетворительное состояние"
    return text


def style_conditions(df, condition_col="Состояние"):
    """Красит строки с неудовлетворительным состоянием красным."""
    if df is None or df.empty:
        return df

    def row_style(row):
        bad = is_bad_condition(row.get(condition_col, ""))
        return [BAD_ROW_STYLE if bad else "" for _ in row]

    return df.style.apply(row_style, axis=1)


def style_violations(df, bad_col="Неудовлетворительные"):
    """Подсвечивает красным колонку с неудовлетворительной техникой."""
    if df is None or df.empty:
        return df

    def row_style(row):
        styles = []
        for col in row.index:
            value = str(row.get(col, ""))
            if col == bad_col and value not in ("", "—"):
                styles.append(
                    "background-color: #ffe1e1; color: #a40000; font-weight: 600"
                )
            else:
                styles.append("")
        return styles

    return df.style.apply(row_style, axis=1)


def laptops_summary(df):
    """Сколько всего ноутбуков, сколько исправных и сколько нет."""
    if df is None or df.empty or "category" not in df:
        return {"total": 0, "ok": 0, "bad": 0}
    lap = df[df["category"] == "Ноутбук"]
    if lap.empty:
        return {"total": 0, "ok": 0, "bad": 0}
    qty = pd.to_numeric(lap["quantity"], errors="coerce").fillna(0).astype(int)
    total = int(qty.sum())
    bad = int(qty[lap["condition"].map(is_bad_condition)].sum())
    return {"total": total, "ok": total - bad, "bad": bad}


def find_limit_violations(df, with_details=True):
    """Партии, где превышен лимит техники. Возвращает DataFrame.

    Колонки: Партия, Категория, Кол-во, Лимит, Превышение [, Позиции].
    Склад «База» из проверки исключён (там техника хранится, а не выдана партии).
    """
    cols = ["Партия", "Категория", "Кол-во", "Лимит", "Превышение"]
    if with_details:
        cols += ["Неудовлетворительные", "Позиции"]
    if df is None or df.empty or "party" not in df:
        return pd.DataFrame(columns=cols)

    work = df[df["party"] != BASE_PARTY].copy()
    if work.empty:
        return pd.DataFrame(columns=cols)
    work["quantity"] = pd.to_numeric(
        work["quantity"], errors="coerce"
    ).fillna(0).astype(int)

    rows = []
    for (party, category), group in work.groupby(["party", "category"]):
        total = int(group["quantity"].sum())
        limit = category_limit(category)
        if total <= limit:
            continue
        row = {
            "Партия": party,
            "Категория": category,
            "Кол-во": total,
            "Лимит": limit,
            "Превышение": total - limit,
        }
        if with_details:
            parts = []
            bad_parts = []
            for _, item in group.iterrows():
                label = str(item.get("model", "")).strip() or category
                serial = str(item.get("serial_number", "")).strip()
                if serial and serial != "-":
                    label += f" (Сер: {serial})"
                qty_text = f"{int(item['quantity'])} шт"
                if is_bad_condition(item.get("condition", "")):
                    bad_info = f"{label} — {condition_text(item.get('condition', ''))}"
                    bad_parts.append(bad_info)
                    parts.append(f"⛔ {bad_info} — {qty_text}")
                else:
                    parts.append(f"{label} — {qty_text}")
            row["Неудовлетворительные"] = (
                "; ".join(bad_parts) if bad_parts else "—"
            )
            row["Позиции"] = "; ".join(parts)
        rows.append(row)

    result = pd.DataFrame(rows, columns=cols)
    if not result.empty:
        result = result.sort_values(["Партия", "Категория"])
    return result


# --- ЖУРНАЛ ДЕЙСТВИЙ (АУДИТ) ---
def log_action(session, action, item, engineer, details=""):
    """Записывает действие (добавление / изменение / удаление) в журнал."""
    session.add(
        AuditLog(
            date=utc_now_str(),
            action=action,
            party=getattr(item, "party", "") or "",
            category=getattr(item, "category", "") or "",
            model=getattr(item, "model", "") or "",
            serial_number=getattr(item, "serial_number", "") or "",
            inv_number=getattr(item, "inv_number", "") or "",
            engineer=engineer or "",
            details=details or "",
        )
    )


def describe_changes(old_values, new_values):
    """Текст «что изменилось» по ключевым полям позиции."""
    labels = {
        "model": "Модель",
        "serial_number": "Серийный номер",
        "inv_number": "Инвентарный номер",
        "quantity": "Количество",
        "condition": "Состояние",
    }
    parts = []
    for key, label in labels.items():
        before = str(old_values.get(key, "") or "")
        after = str(new_values.get(key, "") or "")
        if before != after:
            parts.append(f"{label}: «{before}» → «{after}»")
    return "; ".join(parts) if parts else "без изменений"


# --- QR-КОДЫ И КАРТОЧКА ПОЗИЦИИ ---
def app_base_url():
    """Базовый адрес приложения (нужен для ссылок в QR-кодах)."""
    try:
        host = st.context.headers.get("Host", "")
    except Exception:
        host = ""
    if not host:
        return ""
    scheme = "http" if host.split(":")[0] in ("localhost", "127.0.0.1") else "https"
    return f"{scheme}://{host}"


def item_qr_url(item_id):
    """Ссылка на карточку позиции (то, что зашито в QR-код)."""
    base = app_base_url()
    return f"{base}/?item={item_id}" if base else ""


def qr_svg_bytes(data, box_size=10, border=2):
    """QR-код в векторном SVG (годится для печати наклеек)."""
    img = qrcode.make(
        data,
        image_factory=qrcode.image.svg.SvgPathImage,
        box_size=box_size,
        border=border,
    )
    buf = io.BytesIO()
    img.save(buf)
    return buf.getvalue()


def qr_data_uri(data):
    return (
        "data:image/svg+xml;base64,"
        + base64.b64encode(qr_svg_bytes(data)).decode()
    )


def build_qr_stickers_html(rows, base_url):
    """Страница с сеткой QR-наклеек для печати (A4, печать из браузера)."""
    cards = ""
    for row in rows:
        if not QR_AVAILABLE or not base_url:
            break
        url = f"{base_url}/?item={row['id']}"
        uri = qr_data_uri(url)
        cards += f"""
        <div class="sticker">
          <img src="{uri}" alt="QR {_esc(row.get('id'))}">
          <div class="cap">
            <div class="cap-title">{_esc(row.get('category', ''))} — {_esc(row.get('model', ''))}</div>
            <div class="cap-line">Партия: <b>{_esc(row.get('party', ''))}</b></div>
            <div class="cap-line">Сер: {_esc(row.get('serial_number', ''))} · Инв: {_esc(row.get('inv_number', ''))}</div>
            <div class="cap-id">ID {_esc(row.get('id'))}</div>
          </div>
        </div>"""
    return f"""<!DOCTYPE html>
<html lang="ru"><head><meta charset="utf-8">
<title>QR-наклейки на оргтехнику</title>
<style>
  body {{ font-family: -apple-system, "Segoe UI", Roboto, Arial, sans-serif;
         margin: 16px; color: #111827; }}
  h1 {{ font-size: 18px; }}
  .grid {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; }}
  .sticker {{ display: flex; gap: 12px; align-items: center; border: 1px dashed #9ca3af;
             border-radius: 12px; padding: 10px; page-break-inside: avoid; }}
  .sticker img {{ width: 120px; height: 120px; }}
  .cap-title {{ font-weight: 700; font-size: 14px; }}
  .cap-line {{ font-size: 12px; color: #374151; margin-top: 3px; }}
  .cap-id {{ font-size: 11px; color: #9ca3af; margin-top: 4px; }}
  @media print {{ .sticker {{ border-color: #d1d5db; }} }}
</style></head>
<body>
  <h1>QR-наклейки на оргтехнику</h1>
  <p style="font-size:12px;color:#6b7280">Напечатайте, разрежьте по пунктирным линиям и наклейте на технику.
     Сканирование QR открывает карточку позиции.</p>
  <div class="grid">{cards}</div>
</body></html>"""


def render_item_card(item, session):
    """Карточка позиции — открывается по ссылке из QR-кода (?item=ID)."""
    bad = is_bad_condition(item.condition)
    state_color = "#dc2626" if bad else "#16a34a"
    qr_url = item_qr_url(item.id)

    st.markdown(
        f"""
        <div class="item-card">
          <div class="ic-title">🏷 Карточка позиции</div>
          <h1>{_esc(item.category)} — {_esc(item.model)}</h1>
          <table class="card-table">
            <tr><td>Партия</td><td><b>{_esc(item.party)}</b></td></tr>
            <tr><td>Серийный номер</td><td>{_esc(item.serial_number)}</td></tr>
            <tr><td>Инвентарный номер</td><td>{_esc(item.inv_number)}</td></tr>
            <tr><td>Количество</td><td>{_esc(item.quantity)} шт</td></tr>
            <tr><td>Состояние</td>
                <td style="color:{state_color};font-weight:700">
                    {_esc(item.condition)}</td></tr>
            <tr><td>Ответственный</td><td>{_esc(item.engineer)}</td></tr>
            <tr><td>Обновлено</td><td>{_esc(item.date_updated)}</td></tr>
          </table>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2 = st.columns([1, 2])
    with c1:
        if QR_AVAILABLE and qr_url:
            st.markdown(
                f'<img src="{qr_data_uri(qr_url)}" width="190" height="190" '
                f'alt="QR-код позиции">',
                unsafe_allow_html=True,
            )
        else:
            st.caption("QR-код недоступен.")
    with c2:
        if qr_url:
            st.markdown(
                "**Этот QR-код ведёт на данную карточку.** Его можно "
                "распечатать и наклеить на технику."
            )
            st.code(qr_url, language=None)
        base = app_base_url()
        if base:
            st.markdown(f"[← Вернуться в приложение]({base}/)")

    moves = (
        session.query(History)
        .filter(History.equipment_id == item.id)
        .order_by(History.id.desc())
        .all()
    )
    st.markdown("#### История перемещений по позиции")
    if moves:
        moves_df = pd.DataFrame(
            [
                {
                    "Дата": m.date,
                    "Откуда": m.from_where,
                    "Куда": m.to_where,
                    "Переместил(а)": m.engineer,
                    "Принял(а)": m.receiver,
                }
                for m in moves
            ]
        )
        st.dataframe(moves_df, use_container_width=True, hide_index=True)
    else:
        st.caption("Перемещений по этой позиции пока не было.")


def _esc(value):
    return html_lib.escape(str(value if value is not None else ""))


def _donut(ok, bad, title, size=176):
    """Круговая диаграмма «исправно / неудовлетворительно» в SVG."""
    total = (ok or 0) + (bad or 0)
    r = 62
    c = 2 * 3.141592653589793 * r
    if total <= 0:
        ok_len = 0
        bad_len = 0
        bad_pct = 0
    else:
        ok_len = c * ok / total
        bad_len = c - ok_len
        bad_pct = round(bad * 100 / total)
    ok_pct = 100 - bad_pct if total else 0
    return f"""
    <div class="donut-card">
      <div class="donut-title">{_esc(title)}</div>
      <svg width="{size}" height="{size}" viewBox="0 0 160 160">
        <circle cx="80" cy="80" r="{r}" fill="none"
                stroke="#e5e7eb" stroke-width="18"/>
        <circle cx="80" cy="80" r="{r}" fill="none" stroke="#16a34a"
                stroke-width="18" stroke-dasharray="{ok_len:.2f} {c:.2f}"
                transform="rotate(-90 80 80)" stroke-linecap="butt"/>
        <circle cx="80" cy="80" r="{r}" fill="none" stroke="#dc2626"
                stroke-width="18" stroke-dasharray="{bad_len:.2f} {c:.2f}"
                stroke-dashoffset="{-ok_len:.2f}"
                transform="rotate(-90 80 80)" stroke-linecap="butt"/>
        <text x="80" y="74" text-anchor="middle" font-size="26"
              font-weight="700" fill="#111827">{bad_pct}%</text>
        <text x="80" y="96" text-anchor="middle" font-size="11"
              fill="#6b7280">неисправно</text>
      </svg>
      <div class="donut-legend">
        <span><i class="dot green"></i>Исправно: <b>{ok} ({ok_pct}%)</b></span>
        <span><i class="dot red"></i>Неисправно: <b>{bad} ({bad_pct}%)</b></span>
      </div>
    </div>"""


def _bars(series, title):
    """Горизонтальные бары: series — pandas.Series (значение по названию)."""
    if series is None or len(series) == 0:
        rows = '<div class="muted">Нет данных</div>'
    else:
        top = max(int(v) for v in series.values) or 1
        rows = ""
        for name, value in series.items():
            width = max(round(int(value) * 100 / top), 4)
            rows += (
                f'<div class="bar-row"><div class="bar-name">{_esc(name)}</div>'
                f'<div class="bar-track"><div class="bar-fill" '
                f'style="width:{width}%"></div></div>'
                f'<div class="bar-val">{int(value)}</div></div>'
            )
    return f'<div class="chart-card"><div class="donut-title">{_esc(title)}</div>{rows}</div>'


def _table(headers, rows, empty_text="Нет данных"):
    if not rows:
        return f'<div class="muted">{_esc(empty_text)}</div>'
    head = "".join(f"<th>{_esc(h)}</th>" for h in headers)
    body = ""
    for row in rows:
        cells = ""
        for value in row:
            cls = ""
            text = str(value)
            if text.startswith("!"):  # пометка «плохое состояние»
                cls = ' class="cell-bad"'
                text = text[1:]
            cells += f"<td{cls}>{_esc(text)}</td>"
        body += f"<tr>{cells}</tr>"
    return f'<table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table>'


def build_dashboard_html(df):
    """Собирает автономный HTML-дашборд (открывается в браузере, печатается в PDF)."""
    laptops = laptops_summary(df)
    violations = find_limit_violations(df)
    total_units = int(pd.to_numeric(df["quantity"], errors="coerce").fillna(0).sum()) if not df.empty else 0
    parties_count = int(df["party"].nunique()) if not df.empty else 0
    positions = int(len(df))

    if df.empty:
        defective = df
    else:
        defective = df[df["condition"].map(is_bad_condition)]
    bad_units = int(pd.to_numeric(defective["quantity"], errors="coerce").fillna(0).sum()) if not defective.empty else 0

    by_category = (
        df.groupby("category")["quantity"].sum().sort_values(ascending=False)
        if not df.empty else pd.Series(dtype=int)
    )
    # все позиции без склада — для общего доната «исправно / нет»
    work = df if df.empty else df
    good_units = total_units - bad_units

    kpis = [
        ("📦", "Всего единиц техники", total_units, ""),
        ("🗂", "Позиций в реестре", positions, ""),
        ("🏷", "Партий с техникой", parties_count, ""),
        ("💻", "Ноутбуков всего", laptops["total"], ""),
        ("✅", "Ноутбуки исправны", laptops["ok"], "green"),
        ("⛔", "Ноутбуки неисправны", laptops["bad"], "red"),
        ("⚠️", "Позиций «неудовлетворительно»", len(defective), "red"),
        ("🚫", "Превышений лимита", len(violations), "amber" if len(violations) else "green"),
    ]
    kpi_html = "".join(
        f'<div class="kpi {cls}"><div class="kpi-icon">{icon}</div>'
        f'<div><div class="kpi-value">{value}</div>'
        f'<div class="kpi-label">{_esc(label)}</div></div></div>'
        for icon, label, value, cls in kpis
    )

    # Таблица превышений лимитов
    lim_rows = []
    for _, row in violations.iterrows():
        bad_info = str(row.get("Неудовлетворительные", "—"))
        lim_rows.append([
            row["Партия"], row["Категория"], row["Кол-во"], row["Лимит"],
            row["Превышение"],
            ("!" + bad_info) if bad_info not in ("", "—") else "—",
            row.get("Позиции", ""),
        ])

    # Неудовлетворительные позиции
    def_rows = []
    for _, row in defective.iterrows():
        state = str(row["condition"])
        def_rows.append([
            row["party"], row["category"], row["model"],
            row["serial_number"], row["inv_number"], row["quantity"],
            "!" + state, row["engineer"], row["date_updated"],
        ])

    # Ноутбуки (подробно)
    lap_rows = []
    if not df.empty:
        for _, row in df[df["category"] == "Ноутбук"].iterrows():
            state = str(row["condition"])
            mark = state if not is_bad_condition(state) else "!" + state
            lap_rows.append([
                row["party"], row["model"], row["serial_number"],
                row["inv_number"], mark, row["engineer"],
            ])

    # Сводка по партиям
    if df.empty:
        summary_rows = []
        cat_cols = []
    else:
        pivot = (
            df.groupby(["party", "category"])["quantity"].sum().reset_index()
            .pivot(index="party", columns="category", values="quantity")
            .fillna(0)
        )
        pivot["ИТОГО"] = pivot.sum(axis=1)
        cat_cols = list(pivot.columns)
        summary_rows = []
        for party, row in pivot.iterrows():
            summary_rows.append([party] + [int(row[c]) for c in cat_cols])

    limit_block = (
        _table(["Партия", "Категория", "Кол-во", "Лимит", "Превышение",
                "Неудовлетворительные", "Позиции"],
               lim_rows, "Превышений лимитов нет — всё в норме ✅")
    )
    generated = datetime.now(timezone.utc).strftime("%d.%m.%Y %H:%M UTC")

    return f"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Дашборд учёта оргтехники</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ margin: 0; background: #f5f7fb; color: #111827;
         font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
                      "Helvetica Neue", Arial, sans-serif; }}
  .wrap {{ max-width: 1120px; margin: 0 auto; padding: 28px 20px 60px; }}
  header {{ background: linear-gradient(120deg, #2f6fed, #7c5ce0);
           color: #fff; border-radius: 20px; padding: 28px 30px;
           box-shadow: 0 14px 30px -18px rgba(47,111,237,.9); }}
  header h1 {{ margin: 0; font-size: 26px; letter-spacing: -.02em; }}
  header .sub {{ margin-top: 6px; opacity: .9; font-size: 14px; }}
  .kpis {{ display: grid; gap: 14px; margin: 22px 0;
          grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); }}
  .kpi {{ display: flex; gap: 12px; align-items: center; background: #fff;
         border: 1px solid #e5e7eb; border-radius: 16px; padding: 16px;
         box-shadow: 0 8px 20px -18px rgba(17,24,39,.5); }}
  .kpi-icon {{ font-size: 24px; }}
  .kpi-value {{ font-size: 24px; font-weight: 800; line-height: 1.1; }}
  .kpi-label {{ font-size: 12.5px; color: #6b7280; margin-top: 2px; }}
  .kpi.green .kpi-value {{ color: #16a34a; }}
  .kpi.red .kpi-value {{ color: #dc2626; }}
  .kpi.amber .kpi-value {{ color: #d97706; }}
  .charts {{ display: grid; gap: 16px; margin: 22px 0;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); }}
  .donut-card, .chart-card {{ background: #fff; border: 1px solid #e5e7eb;
      border-radius: 18px; padding: 18px; box-shadow: 0 8px 20px -18px rgba(17,24,39,.5); }}
  .donut-title {{ font-weight: 700; margin-bottom: 12px; font-size: 15px; }}
  .donut-card {{ display: flex; flex-direction: column; align-items: center; }}
  .donut-legend {{ display: flex; gap: 16px; flex-wrap: wrap; font-size: 13px;
                  margin-top: 8px; color: #374151; }}
  .dot {{ display: inline-block; width: 10px; height: 10px; border-radius: 50%;
         margin-right: 5px; }}
  .dot.green {{ background: #16a34a; }}
  .dot.red {{ background: #dc2626; }}
  .bar-row {{ display: grid; grid-template-columns: 190px 1fr 46px; gap: 10px;
             align-items: center; margin: 9px 0; font-size: 13.5px; }}
  .bar-name {{ color: #374151; }}
  .bar-track {{ background: #eef2f7; border-radius: 999px; height: 12px;
               overflow: hidden; }}
  .bar-fill {{ height: 100%; border-radius: 999px;
              background: linear-gradient(90deg, #2f6fed, #7c5ce0); }}
  .bar-val {{ text-align: right; font-weight: 700; }}
  h2 {{ font-size: 18px; margin: 30px 0 12px; }}
  .panel {{ background: #fff; border: 1px solid #e5e7eb; border-radius: 18px;
           padding: 18px; box-shadow: 0 8px 20px -18px rgba(17,24,39,.5);
           overflow-x: auto; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 13.5px; }}
  th {{ text-align: left; background: #f1f5f9; color: #334155; font-weight: 700;
       padding: 10px 12px; border-bottom: 1px solid #e5e7eb; white-space: nowrap; }}
  td {{ padding: 9px 12px; border-bottom: 1px solid #eef2f7; vertical-align: top; }}
  tr:last-child td {{ border-bottom: none; }}
  tbody tr:hover {{ background: #f8fafc; }}
  .cell-bad {{ color: #a40000; font-weight: 700; background: #ffe1e1; }}
  .muted {{ color: #6b7280; font-size: 14px; padding: 8px 0; }}
  .note {{ font-size: 12.5px; color: #6b7280; margin-top: 8px; }}
  @media print {{
    body {{ background: #fff; }}
    .wrap {{ max-width: none; padding: 0; }}
    .kpi, .panel, .donut-card, .chart-card, header {{ box-shadow: none; }}
  }}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>💻 Дашборд учёта оргтехники</h1>
    <div class="sub">Полевые инженеры · сформирован {generated}</div>
  </header>

  <section class="kpis">{kpi_html}</section>

  <section class="charts">
    {_donut(good_units, bad_units, "Вся техника: исправно / неисправно")}
    {_donut(laptops["ok"], laptops["bad"], "Ноутбуки: исправно / неисправно")}
    {_bars(by_category, "Количество по категориям")}
  </section>

  <h2>🚫 Партии с превышением лимита</h2>
  <div class="panel">{limit_block}
    <div class="note">Лимит: ноутбуки и сетевые фильтры — не более 2 шт в партии,
      остальная оргтехника — не более 1 шт. Склад «База» не проверяется.</div>
  </div>

  <h2>⛔ Позиции в неудовлетворительном состоянии</h2>
  <div class="panel">{_table(
      ["Партия", "Категория", "Модель", "Серийный №", "Инвентарный №",
       "Кол-во", "Состояние", "Ответственный", "Обновлено"],
      def_rows, "Таких позиций нет ✅")}</div>

  <h2>💻 Ноутбуки</h2>
  <div class="panel">{_table(
      ["Партия", "Модель", "Серийный №", "Инвентарный №", "Состояние",
       "Ответственный"], lap_rows, "Ноутбуков нет")}</div>

  <h2>📊 Сводка по партиям</h2>
  <div class="panel">{_table(["Партия"] + cat_cols, summary_rows,
                              "Данных нет")}</div>
</div>
</body>
</html>"""


# --- НАСТРОЙКА БАЗЫ ДАННЫХ ---
# Приоритет: 1) переменная окружения DATABASE_URL; 2) секрет database_url
# (Streamlit Cloud: App → Settings → Secrets) — постоянный Postgres;
# 3) локальный SQLite-файл (запасной режим).
DB_FILE = "inventory_v4.db"
DATABASE_URL = os.environ.get("DATABASE_URL", "") or st.secrets.get(
    "database_url", ""
)
if DATABASE_URL:
    engine = create_engine(DATABASE_URL, echo=False, pool_pre_ping=True)
    IS_SQLITE = False
else:
    engine = create_engine(f"sqlite:///{DB_FILE}", echo=False)
    IS_SQLITE = True
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()


class Equipment(Base):
    __tablename__ = "equipment"
    id = Column(Integer, primary_key=True, autoincrement=True)
    party = Column(String)
    category = Column(String)
    model = Column(String)
    serial_number = Column(String)
    inv_number = Column(String)
    quantity = Column(Integer)
    condition = Column(String)
    engineer = Column(String)
    date_updated = Column(String)


class LaptopReference(Base):
    __tablename__ = "laptop_reference"
    id = Column(Integer, primary_key=True, autoincrement=True)
    model = Column(String)
    serial_number = Column(String)
    inv_number = Column(String)


class CustomModels(Base):
    __tablename__ = "custom_models"
    id = Column(Integer, primary_key=True, autoincrement=True)
    category = Column(String)
    model = Column(String)


class History(Base):
    __tablename__ = "history"
    id = Column(Integer, primary_key=True, autoincrement=True)
    date = Column(String)
    equipment_info = Column(String)
    from_where = Column(String)
    to_where = Column(String)
    engineer = Column(String)
    receiver = Column(String)
    equipment_id = Column(Integer)


class AuditLog(Base):
    """Журнал действий: кто и когда добавил, изменил или удалил позицию."""

    __tablename__ = "audit_log"
    id = Column(Integer, primary_key=True, autoincrement=True)
    date = Column(String)
    action = Column(String)
    party = Column(String)
    category = Column(String)
    model = Column(String)
    serial_number = Column(String)
    inv_number = Column(String)
    engineer = Column(String)
    details = Column(String)


Base.metadata.create_all(bind=engine)


def migrate_history_schema():
    """Добавляет новые колонки в существующую таблицу истории без потери данных.
    Работает и в SQLite, и в Postgres (через SQLAlchemy Inspector)."""
    if not inspect(engine).has_table("history"):
        return
    existing = {c["name"] for c in inspect(engine).get_columns("history")}
    with engine.begin() as conn:
        if "receiver" not in existing:
            conn.execute(text("ALTER TABLE history ADD COLUMN receiver VARCHAR"))
        if "equipment_id" not in existing:
            conn.execute(
                text("ALTER TABLE history ADD COLUMN equipment_id INTEGER")
            )


migrate_history_schema()


# Автоматическая загрузка ноутбуков из Excel
def init_laptop_db():
    session = SessionLocal()
    if session.query(LaptopReference).count() == 0:
        file_path = "Ноутбуки перемещения 29.08.2024 г..xlsx"
        if os.path.exists(file_path):
            try:
                df = pd.read_excel(file_path, sheet_name="Ноутбуки")
                for _, row in df.iterrows():
                    model = str(row.get("Модель", ""))
                    if model == "nan" or not model:
                        continue

                    condition_text = str(
                        row.get("Состояние", "")
                    ).upper() + str(row.get("Обьект заказчика", "")).upper()
                    if (
                        "СПИСАТЬ" in condition_text
                        or "У СИСАДМИНА НА СПИСАНИЕ" in condition_text
                    ):
                        continue

                    serial = str(row.get("Завадской серийный номер", ""))
                    if serial == "nan" or not serial:
                        serial = "-"

                    inv = str(row.get("Инвентарный номер", ""))
                    if inv == "nan" or not inv:
                        inv = "-"

                    ref_item = LaptopReference(
                        model=model, serial_number=serial, inv_number=inv
                    )
                    session.add(ref_item)
                session.commit()
            except Exception as e:
                print(f"Ошибка импорта ноутбуков: {e}")
    session.close()


init_laptop_db()

# --- ИНТЕРФЕЙС STREAMLIT ---
st.set_page_config(
    page_title="Учет оргтехники", page_icon="💻", layout="centered"
)

# --- ЕДИНЫЙ СТИЛЬ ИНТЕРФЕЙСА ---
# Цвета берутся из переменных темы Streamlit (var(--...)),
# поэтому приложение одинаково хорошо выглядит в светлой и тёмной теме.
st.markdown(
    """
    <style>
    /* Базовые шрифты и отступы */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto,
            "Helvetica Neue", Arial, sans-serif;
    }
    [data-testid="stAppViewContainer"] {
        background: var(--background-color);
    }
    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 4rem;
        max-width: 960px;
    }

    /* Акцентная полоса сверху */
    [data-testid="stAppViewContainer"]::before {
        content: "";
        position: fixed;
        top: 0; left: 0; right: 0;
        height: 4px;
        z-index: 99999;
        background: linear-gradient(90deg, #2f6fed, #7c5ce0, #2f6fed);
    }

    /* Заголовки */
    h1, h2, h3 { letter-spacing: -0.01em; }
    h1 { font-weight: 800 !important; }
    h3 { font-weight: 700 !important; }

    /* Шапка приложения (hero) */
    .hero {
        display: flex;
        align-items: center;
        gap: 1rem;
        padding: 1.1rem 1.25rem;
        margin-bottom: 1.2rem;
        background: var(--secondary-background-color);
        border: 1px solid rgba(128, 128, 128, 0.18);
        border-radius: 18px;
    }
    .hero-icon {
        font-size: 2.1rem;
        line-height: 1;
        flex-shrink: 0;
    }
    .hero h1 {
        margin: 0;
        font-size: 1.5rem;
        color: var(--text-color);
    }
    .hero-sub {
        margin-top: 0.15rem;
        font-size: 0.85rem;
        color: var(--text-color);
        opacity: 0.65;
    }

    /* Бейдж партии (цвет задан явно — не зависит от темы Streamlit) */
    .party-chip {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        padding: 0.3rem 0.9rem;
        border-radius: 999px;
        font-weight: 700;
        font-size: 0.95rem;
        color: #2f6fed;
        background: rgba(47, 111, 237, 0.12);
        border: 1px solid rgba(47, 111, 237, 0.32);
        margin-bottom: 0.6rem;
    }

    /* Основные кнопки и активные вкладки — фирменный синий */
    button[kind="primary"], button[data-testid="stBaseButton-primary"] {
        background-color: #2f6fed !important;
    }
    button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover {
        background-color: #285fd0 !important;
    }
    [data-testid="stTabs"] button[data-baseweb="tab"][aria-selected="true"] {
        color: #2f6fed !important;
        border-bottom: 2px solid #2f6fed !important;
    }

    /* Кнопки */
    .stButton button, .stDownloadButton button, .stFormSubmitButton button {
        border-radius: 12px;
        font-weight: 600;
        transition: transform 0.12s ease, box-shadow 0.12s ease;
    }
    .stButton button:hover, .stDownloadButton button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px -8px rgba(0, 0, 0, 0.35);
    }
    .stButton button:active, .stDownloadButton button:active {
        transform: translateY(0);
    }

    /* Подсказка только для мобильных (меню спрятано за гамбургером) */
    .device-hint {
        display: none;
    }

    /* Карточки статистики по ноутбукам */
    .laptop-stats {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 0.75rem;
        margin-bottom: 0.5rem;
    }
    .ls-card {
        background: var(--secondary-background-color);
        border: 1px solid rgba(128, 128, 128, 0.18);
        border-radius: 14px;
        padding: 0.9rem 1rem;
        text-align: center;
    }
    .ls-val { font-size: 1.7rem; font-weight: 800; line-height: 1.1; }
    .ls-label { font-size: 0.8rem; opacity: 0.7; margin-top: 0.2rem; }
    .ls-card.ok .ls-val { color: #16a34a; }
    .ls-card.bad .ls-val { color: #dc2626; }

    /* Карточка позиции (открывается по QR-ссылке) */
    .item-card {
        background: var(--secondary-background-color);
        border: 1px solid rgba(128, 128, 128, 0.18);
        border-radius: 18px;
        padding: 1.2rem 1.4rem;
        margin-bottom: 1rem;
    }
    .item-card .ic-title {
        font-size: 0.85rem;
        opacity: 0.65;
        margin-bottom: 0.3rem;
    }
    .item-card h1 { margin: 0 0 0.8rem; font-size: 1.4rem; }
    .card-table { width: 100%; border-collapse: collapse; }
    .card-table td {
        padding: 0.35rem 0.2rem;
        border-bottom: 1px solid rgba(128, 128, 128, 0.12);
        font-size: 0.92rem;
        vertical-align: top;
    }
    .card-table td:first-child { color: #6b7280; width: 42%; }

    /* Вкладки */
    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        gap: 0.25rem;
    }
    [data-testid="stTabs"] button[data-baseweb="tab"] {
        font-weight: 600;
        border-radius: 10px 10px 0 0;
        padding: 0.55rem 1rem;
    }

    /* Скрываем служебный футер Streamlit */
    [data-testid="stFooter"] { display: none; }

    /* ---------- МОБИЛЬНАЯ ВЕРСИЯ ---------- */
    @media (max-width: 768px) {
        .hero { padding: 0.9rem 1rem; border-radius: 14px; }
        .hero h1 { font-size: 1.25rem; }

        .device-hint {
            display: block;
            padding: 0.75rem 0.9rem;
            margin-bottom: 1rem;
            background: #fff4d6;
            color: #5c4400;
            border: 1px solid #ffd54f;
            border-radius: 12px;
            font-size: 0.9rem;
            line-height: 1.45;
        }

        /* Все поля и кнопки — 16px, чтобы iOS не увеличивал масштаб */
        .stButton button, .stDownloadButton button {
            width: 100%;
            font-size: 16px;
            padding: 0.65rem 0.5rem;
        }
        input, textarea, select, [data-baseweb="select"] > div {
            font-size: 16px !important;
        }
        [data-testid="stTabs"] button[data-baseweb="tab"] {
            font-size: 0.85rem;
            padding: 0.45rem 0.6rem;
        }
        .laptop-stats { gap: 0.5rem; }
        .ls-card { padding: 0.7rem 0.5rem; border-radius: 12px; }
        .ls-val { font-size: 1.4rem; }
        .ls-label { font-size: 0.7rem; }
        .block-container { padding-top: 0.9rem; }
    }
    </style>
""",
    unsafe_allow_html=True,
)

# --- ШАПКА ПРИЛОЖЕНИЯ ---
st.markdown(
    """
    <div class="hero">
        <div class="hero-icon">💻</div>
        <div>
            <h1>Учёт оргтехники полевых инженеров</h1>
            <div class="hero-sub">Инвентаризация, партии и перемещения техники</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

parties = [f"Партия № {i}" for i in range(1, 32)]

cats_with_identifiers = [
    "Ноутбук",
    "Моноблок",
    "Роутер Huawei",
    "Коммутатор/хаб для обменника",
    "МФУ/Принтер",
    "Усилитель сотовой связи",
    "Сотовый телефон",
    "Стабилизатор напряжения",
    "ИБП",
]

cats_with_qty = [
    "ПК/Системный блок",
    "Монитор",
    "Сетевой фильтр",
    "Lan кабель 50 метров",
]

session = SessionLocal()

# --- ОТКРЫТИЕ ПО QR-КОДУ: карточка позиции (?item=ID) ---
qr_item_param = (st.query_params.get_all("item") or [""])[0]
if qr_item_param.isdigit():
    card_item = (
        session.query(Equipment)
        .filter(Equipment.id == int(qr_item_param))
        .first()
    )
    if card_item:
        render_item_card(card_item, session)
    else:
        st.error("Позиция не найдена — возможно, она была удалена.")
    session.close()
    st.stop()

# --- БОКОВАЯ ПАНЕЛЬ НАВИГАЦИИ ---
st.sidebar.header("Параметры сеанса")
role = st.sidebar.selectbox("Режим работы:", ["Инженер", "Администратор"])

current_engineer = ""
selected_party = ""

if role == "Инженер":
    # Восстановление партии и ФИО из адресной строки (?party=...&fio=...):
    # инженер, вернувшийся по закладке/истории браузера, попадает сразу в свою
    # партию с уже заполненным ФИО.
    q_party = (st.query_params.get_all("party") or [""])[0]
    q_fio = (st.query_params.get_all("fio") or [""])[0]

    party_index = parties.index(q_party) if q_party in parties else 0
    selected_party = st.sidebar.selectbox(
        "Выберите вашу партию:", parties, index=party_index
    )

    if (
        "last_party" not in st.session_state
        or st.session_state["last_party"] != selected_party
    ):
        st.session_state["last_party"] = selected_party
        # ФИО подставляем только если в URL та же партия (тот же контекст)
        st.session_state["engineer_name"] = (
            q_fio if q_party == selected_party else ""
        )

    current_engineer = st.sidebar.text_input(
        "ФИО ответственного инженера:",
        value=st.session_state["engineer_name"],
        key="eng_input_field",
    )
    st.session_state["engineer_name"] = current_engineer

    # Сохраняем выбор в адресную строку, чтобы он пережил перезагрузку страницы
    if (st.query_params.get_all("party") or [""])[0] != selected_party:
        st.query_params["party"] = selected_party
    url_fio = (st.query_params.get_all("fio") or [""])[0]
    if current_engineer and url_fio != current_engineer:
        st.query_params["fio"] = current_engineer
    elif not current_engineer and "fio" in st.query_params:
        del st.query_params["fio"]

if role == "Инженер":
    st.markdown(
        f'<span class="party-chip">📍 {selected_party}</span>',
        unsafe_allow_html=True,
    )

    if not current_engineer:
        st.markdown(
            '<div class="device-hint">⚠️ <b>Внимание!</b> Нажмите на стрелочку '
            '<b>&gt; (меню)</b> в верхнем левом углу экрана и укажите ваше '
            '<b>ФИО</b> для продолжения работы!</div>',
            unsafe_allow_html=True,
        )
    else:
        tab1, tab2, tab3 = st.tabs(
            ["📋 Список", "➕ Добавить", "🚚 Переместить"]
        )

        with tab1:
            data = pd.read_sql(
                session.query(Equipment)
                .filter(Equipment.party == selected_party)
                .statement,
                engine,
            )
            if not data.empty:
                st.dataframe(
                    style_conditions(
                        data[
                            [
                                "category",
                                "model",
                                "serial_number",
                                "inv_number",
                                "quantity",
                                "condition",
                                "engineer",
                                "date_updated",
                            ]
                        ].rename(columns=COL_RU)
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
                st.caption(
                    "Строки с неудовлетворительным состоянием выделены красным."
                )

                # --- Изменение и удаление позиции ---
                st.markdown("---")
                st.markdown("#### ✏️ Изменить или удалить позицию")
                st.caption(
                    "Выберите позицию: можно исправить данные (модель, номера, "
                    "количество, состояние) или удалить ошибочную запись."
                )
                item_options = {
                    f"{row.category} — {row.model} "
                    f"(Сер: {row.serial_number} | Инв: {row.inv_number})": row.id
                    for _, row in data.iterrows()
                }
                selected_label = st.selectbox(
                    "Выберите позицию", list(item_options.keys())
                )
                item_id = item_options[selected_label]
                item = (
                    session.query(Equipment)
                    .filter(Equipment.id == item_id)
                    .first()
                )
                if item:
                    cur_bad = is_bad_condition(item.condition)
                    st.markdown(f"**Категория:** {item.category} (не изменяется)")
                    with st.form(f"edit_form_{item.id}"):
                        new_model = st.text_input(
                            "Модель", value=item.model or ""
                        )
                        new_serial = st.text_input(
                            "Серийный номер", value=item.serial_number or ""
                        )
                        new_inv = st.text_input(
                            "Инвентарный номер", value=item.inv_number or ""
                        )
                        new_qty = st.number_input(
                            "Количество (шт.)",
                            min_value=1,
                            value=int(item.quantity or 1),
                        )
                        new_cond = st.selectbox(
                            "Состояние",
                            ["Удовлетворительно", "Не удовлетворительно"],
                            index=1 if cur_bad else 0,
                        )
                        new_comment = st.text_input(
                            "Описание неисправности (только для «Не удовлетворительно»)",
                            value=item.condition if cur_bad else "",
                        )
                        save_clicked = st.form_submit_button(
                            "💾 Сохранить изменения", use_container_width=True
                        )
                        st.markdown("**Удаление позиции**")
                        confirm_del = st.checkbox("Подтверждаю удаление позиции")
                        delete_clicked = st.form_submit_button(
                            "🗑️ Удалить позицию", use_container_width=True
                        )

                    if save_clicked:
                        err = None
                        if new_cond == "Не удовлетворительно":
                            if not new_comment.strip():
                                err = (
                                    "Опишите неисправность — это поле обязательно."
                                )
                            final_condition = new_comment.strip()
                        else:
                            final_condition = "Удовлетворительно"

                        dup = None
                        if new_serial.strip() and new_serial.strip() != "-":
                            dup = (
                                session.query(Equipment)
                                .filter(
                                    Equipment.serial_number
                                    == new_serial.strip(),
                                    Equipment.id != item.id,
                                )
                                .first()
                            )
                        if err is None and dup is not None:
                            err = (
                                f"Серийный номер «{new_serial.strip()}» уже "
                                f"используется в партии «{dup.party}»."
                            )
                        has_data = bool(
                            new_model.strip()
                            or (new_serial.strip() and new_serial.strip() != "-")
                            or (new_inv.strip() and new_inv.strip() != "-")
                        )
                        if err is None and not has_data:
                            err = (
                                "Заполните хотя бы одно поле: модель, серийный "
                                "или инвентарный номер."
                            )

                        if err:
                            st.error("❌ " + err)
                        else:
                            old_values = {
                                "model": item.model,
                                "serial_number": item.serial_number,
                                "inv_number": item.inv_number,
                                "quantity": item.quantity,
                                "condition": item.condition,
                            }
                            item.model = new_model.strip() or item.model
                            item.serial_number = new_serial.strip() or "-"
                            item.inv_number = new_inv.strip() or "-"
                            item.quantity = int(new_qty)
                            item.condition = final_condition
                            item.engineer = current_engineer
                            item.date_updated = utc_now_str()
                            new_values = {
                                "model": item.model,
                                "serial_number": item.serial_number,
                                "inv_number": item.inv_number,
                                "quantity": item.quantity,
                                "condition": item.condition,
                            }
                            log_action(
                                session,
                                "Изменение",
                                item,
                                current_engineer,
                                describe_changes(old_values, new_values),
                            )
                            session.commit()
                            st.success("Изменения сохранены!")
                            st.rerun()

                    if delete_clicked:
                        if not confirm_del:
                            st.error(
                                "❌ Отметьте «Подтверждаю удаление», чтобы "
                                "удалить позицию."
                            )
                        else:
                            log_action(
                                session,
                                "Удаление",
                                item,
                                current_engineer,
                                "Позиция удалена инженером",
                            )
                            session.query(Equipment).filter(
                                Equipment.id == item.id
                            ).delete()
                            session.commit()
                            st.success("Позиция удалена!")
                            st.rerun()

                    with st.expander("🏷 QR-код этой позиции"):
                        qr_url = item_qr_url(item.id)
                        if QR_AVAILABLE and qr_url:
                            st.markdown(
                                f'<img src="{qr_data_uri(qr_url)}" '
                                f'width="190" height="190" '
                                f'alt="QR-код позиции">',
                                unsafe_allow_html=True,
                            )
                            st.caption(
                                "Сканирование ведёт на карточку позиции — "
                                "можно распечатать и наклеить на технику."
                            )
                        else:
                            st.caption(
                                "QR-код недоступен: не установлена библиотека "
                                "qrcode или не определён адрес приложения."
                            )
            else:
                st.info("У данной партии пока нет зарегистрированной техники.")

        with tab2:
            st.markdown("### Добавить оргтехнику")
            all_categories = cats_with_identifiers + cats_with_qty
            cat = st.selectbox("Категория техники", all_categories)

            # Выбор состояния и обязательный комментарий при неудовлетворительном
            cond_option = st.selectbox(
                "Состояние", ["Удовлетворительно", "Не удовлетворительно"]
            )
            if cond_option == "Не удовлетворительно":
                cond_comment = st.text_input(
                    "⚠️ Опишите неисправность (обязательно):"
                )
            else:
                cond_comment = "Удовлетворительно"

            # Логика для техники с серийными/инвентарными номерами
            if cat in cats_with_identifiers:
                model, serial, inv = "", "", ""

                if cat == "Ноутбук":
                    laptops_db = session.query(LaptopReference).all()
                    inv_list = [
                        l.inv_number for l in laptops_db if l.inv_number != "-"
                    ]
                    ser_list = [
                        l.serial_number
                        for l in laptops_db
                        if l.serial_number != "-"
                    ]

                    sel_inv = st.selectbox(
                        "Поиск по инвентарному номеру (из базы)", [""] + inv_list
                    )
                    sel_ser = st.selectbox(
                        "Или поиск по серийному номеру (из базы)", [""] + ser_list
                    )

                    auto_model, auto_serial, auto_inv = "", "", ""
                    if sel_inv:
                        match = (
                            session.query(LaptopReference)
                            .filter(LaptopReference.inv_number == sel_inv)
                            .first()
                        )
                        if match:
                            auto_model, auto_serial, auto_inv = (
                                match.model,
                                match.serial_number,
                                match.inv_number,
                            )
                    elif sel_ser:
                        match = (
                            session.query(LaptopReference)
                            .filter(LaptopReference.serial_number == sel_ser)
                            .first()
                        )
                        if match:
                            auto_model, auto_serial, auto_inv = (
                                match.model,
                                match.serial_number,
                                match.inv_number,
                            )

                    model = st.text_input(
                        "Модель ноутбука",
                        value=auto_model,
                        placeholder="Введите или выберите выше",
                    )
                    serial = st.text_input("Серийный номер", value=auto_serial)
                    inv = st.text_input("Инвентарный номер", value=auto_inv)

                elif cat == "Роутер Huawei":
                    model = "Роутер Huawei"
                    serial = st.text_input("Серийный номер")
                    inv = st.text_input("Инвентарный номер")

                else:
                    custom_db = (
                        session.query(CustomModels)
                        .filter(CustomModels.category == cat)
                        .all()
                    )
                    saved_models = [m.model for m in custom_db]

                    model = st.selectbox(
                        "Модель (из сохраненных или новая)",
                        [""] + saved_models,
                    )
                    if not model:
                        model = st.text_input("Или введите модель вручную:")

                    serial = st.text_input("Серийный номер")
                    inv = st.text_input("Инвентарный номер")

                if st.button("Сохранить позицию", use_container_width=True):
                    if (
                        cond_option == "Не удовлетворительно"
                        and not cond_comment.strip()
                    ):
                        st.error(
                            "❌ Ошибка: Обязательно укажите описание неисправности!"
                        )
                    else:
                        is_model_valid = bool(
                            model
                            and model != "Роутер Huawei"
                            and model != "Не указана"
                        )
                        is_serial_valid = bool(serial and serial != "-")
                        is_inv_valid = bool(inv and inv != "-")

                        if not (
                            is_model_valid or is_serial_valid or is_inv_valid
                        ):
                            st.error(
                                "❌ Ошибка: Заполните хотя бы одно поле (Модель, Серийный или Инвентарный номер)!"
                            )
                        else:
                            # Защита от дублей: один серийный номер — одна позиция
                            duplicate_item = None
                            if serial and serial != "-":
                                duplicate_item = (
                                    session.query(Equipment)
                                    .filter(Equipment.serial_number == serial)
                                    .first()
                                )
                            if duplicate_item:
                                st.error(
                                    f"❌ Позиция с серийным номером «{serial}» уже зарегистрирована в партии «{duplicate_item.party}». Дубликат не сохранён."
                                )
                            else:
                                if (
                                    cat not in ["Ноутбук", "Роутер Huawei"]
                                    and model
                                ):
                                    existing = (
                                        session.query(CustomModels)
                                        .filter(
                                            CustomModels.category == cat,
                                            CustomModels.model == model,
                                        )
                                        .first()
                                    )
                                    if not existing:
                                        session.add(
                                            CustomModels(
                                                category=cat, model=model
                                            )
                                        )
                                        session.commit()

                                new_item = Equipment(
                                    party=selected_party,
                                    category=cat,
                                    model=model
                                    if model
                                    else "Роутер Huawei"
                                    if cat == "Роутер Huawei"
                                    else "Не указана",
                                    serial_number=serial if serial else "-",
                                    inv_number=inv if inv else "-",
                                    quantity=1,
                                    condition=cond_comment,
                                    engineer=current_engineer,
                                    date_updated=utc_now_str(),
                                )
                                session.add(new_item)
                                log_action(
                                    session,
                                    "Добавление",
                                    new_item,
                                    current_engineer,
                                    "Позиция добавлена инженером",
                                )
                                session.commit()
                                st.success("Успешно добавлено!")
                                st.rerun()

            else:
                qty = st.number_input("Количество (шт.)", min_value=1, value=1)

                if st.button(
                    "Сохранить количество", use_container_width=True
                ):
                    if (
                        cond_option == "Не удовлетворительно"
                        and not cond_comment.strip()
                    ):
                        st.error(
                            "❌ Ошибка: Обязательно укажите описание неисправности!"
                        )
                    else:
                        new_item = Equipment(
                            party=selected_party,
                            category=cat,
                            model=cat,
                            serial_number="-",
                            inv_number="-",
                            quantity=qty,
                            condition=cond_comment,
                            engineer=current_engineer,
                            date_updated=utc_now_str(),
                        )
                        session.add(new_item)
                        log_action(
                            session,
                            "Добавление",
                            new_item,
                            current_engineer,
                            "Позиция добавлена инженером",
                        )
                        session.commit()
                        st.success("Успешно добавлено!")
                        st.rerun()

        with tab3:
            st.markdown("### Перемещение техники")
            data = pd.read_sql(
                session.query(Equipment)
                .filter(Equipment.party == selected_party)
                .statement,
                engine,
            )
            if not data.empty:
                item_options = {
                    f"{row.category} — {row.model} (Сер: {row.serial_number})": row.id
                    for index, row in data.iterrows()
                }
                selected_item_label = st.selectbox(
                    "Выберите позицию", list(item_options.keys())
                )
                destination = st.selectbox(
                    "Куда переместить?", ["База"] + parties
                )

                st.markdown("#### Подтверждение ответственных")
                st.caption(
                    "Оба поля обязательны: перемещение фиксируется на вас и на принимающего."
                )
                confirm_surname = st.text_input(
                    "Ваша фамилия (кто выполняет перемещение)",
                    placeholder="Например: Иванов",
                )
                receiver_surname = st.text_input(
                    "Фамилия принимающего (кому передаёте технику)",
                    placeholder="Например: Петров",
                )

                if st.button(
                    "Подтвердить перемещение", use_container_width=True
                ):
                    ok, err = validate_transfer_surnames(
                        confirm_surname, receiver_surname, current_engineer
                    )
                    if not ok:
                        st.error("❌ " + err)
                    else:
                        item_id = item_options[selected_item_label]
                        item_to_move = (
                            session.query(Equipment)
                            .filter(Equipment.id == item_id)
                            .first()
                        )
                        if item_to_move:
                            if item_to_move.party == destination:
                                st.error(
                                    "❌ Техника уже находится в этой партии. Выберите другое место."
                                )
                            else:
                                old_party = item_to_move.party
                                item_to_move.party = destination
                                item_to_move.date_updated = utc_now_str()

                                history_entry = History(
                                    date=utc_now_str(),
                                    equipment_info=f"{item_to_move.category} {item_to_move.model}",
                                    from_where=old_party,
                                    to_where=destination,
                                    engineer=current_engineer,
                                    receiver=receiver_surname.strip(),
                                    equipment_id=item_to_move.id,
                                )
                                session.add(history_entry)
                                session.commit()
                                st.success("Перемещение выполнено успешно!")
                                st.rerun()
            else:
                st.info("Нет доступной техники для перемещения.")

elif role == "Администратор":
    st.sidebar.subheader("🔒 Авторизация администратора")

    # Пароль хранится в секретах Streamlit (App → Settings → Secrets → admin_password),
    # а не в коде. Если секрет не настроен — вход невозможен.
    ADMIN_PASSWORD = st.secrets.get("admin_password", "")

    if not ADMIN_PASSWORD:
        st.error(
            "Пароль администратора не настроен. Добавьте в секреты приложения "
            "(`App → Settings → Secrets`) ключ `admin_password`, затем обновите страницу."
        )
        st.stop()

    password = st.sidebar.text_input(
        "Введите пароль администратора", type="password"
    )

    if password == ADMIN_PASSWORD:
        st.success("Добро пожаловать в панель администратора!")
        all_data = pd.read_sql(session.query(Equipment).statement, engine)
        history_data = pd.read_sql(session.query(History).statement, engine)
        audit_data = pd.read_sql(session.query(AuditLog).statement, engine)

        tab_dash, tab_all, tab_hist, tab_move = st.tabs(
            ["📊 Сводка", "📁 Реестр", "📜 История", "🚚 Переместить"]
        )

        with tab_dash:
            if all_data.empty:
                st.warning("В базе пока нет записей.")
            else:
                total_units = int(all_data["quantity"].sum())
                defective = all_data[
                    all_data["condition"].map(is_bad_condition)
                ]
                violations = find_limit_violations(all_data)
                lap = laptops_summary(all_data)

                r1c1, r1c2 = st.columns(2)
                r2c1, r2c2 = st.columns(2)
                r1c1.metric("Всего единиц техники", total_units)
                r1c2.metric(
                    "Позиций «неудовлетворительно»", len(defective)
                )
                r2c1.metric("Партий с техникой", all_data["party"].nunique())
                r2c2.metric("Превышений лимита", len(violations))

                # --- Ноутбуки: всего / исправны / неисправны ---
                st.markdown("### 💻 Ноутбуки")
                st.markdown(
                    f"""
                    <div class="laptop-stats">
                      <div class="ls-card">
                        <div class="ls-val">{lap["total"]}</div>
                        <div class="ls-label">Всего ноутбуков</div>
                      </div>
                      <div class="ls-card ok">
                        <div class="ls-val">{lap["ok"]}</div>
                        <div class="ls-label">Исправны</div>
                      </div>
                      <div class="ls-card bad">
                        <div class="ls-val">{lap["bad"]}</div>
                        <div class="ls-label">Неисправны</div>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

                # --- Превышение лимитов по партиям ---
                st.markdown("### 🚫 Партии с превышением лимита")
                st.caption(
                    "Лимит: ноутбуки и сетевые фильтры — не более 2 шт в партии, "
                    "остальная оргтехника — не более 1 шт. Склад «База» не проверяется."
                )
                if violations.empty:
                    st.success("✅ Превышений лимитов нет.")
                else:
                    if (violations["Неудовлетворительные"] != "—").any():
                        st.warning(
                            "⚠️ В партиях с превышением лимита есть техника в "
                            "неудовлетворительном состоянии — смотрите колонку "
                            "«Неудовлетворительные»."
                        )
                    st.dataframe(
                        style_violations(violations),
                        use_container_width=True,
                        hide_index=True,
                    )

                # --- Неудовлетворительные позиции ---
                st.markdown(
                    "### ⛔ Позиции в неудовлетворительном состоянии"
                )
                if defective.empty:
                    st.success("✅ Таких позиций нет.")
                else:
                    def_cols = [
                        "party", "category", "model", "serial_number",
                        "inv_number", "quantity", "condition", "engineer",
                        "date_updated",
                    ]
                    st.dataframe(
                        style_conditions(
                            defective[def_cols].rename(columns=COL_RU)
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )

                # --- Сводная таблица по партиям ---
                st.markdown("### 📊 Сводка по партиям")
                summary = (
                    all_data.groupby(["party", "category"])["quantity"]
                    .sum()
                    .reset_index()
                )
                pivot_summary = summary.pivot(
                    index="party", columns="category", values="quantity"
                ).fillna(0)
                pivot_summary.loc["ИТОГО по категориям"] = pivot_summary.sum(axis=0)
                pivot_summary["ИТОГО по партии"] = pivot_summary.sum(axis=1)
                st.dataframe(pivot_summary, use_container_width=True)

                # --- Скачивание отчётов ---
                st.markdown("---")
                st.markdown("### 📥 Выгрузка отчётов")
                dl1, dl2 = st.columns(2)
                dashboard_html = build_dashboard_html(all_data)
                with dl1:
                    st.download_button(
                        label="📊 Дашборд с графиками (HTML)",
                        data=dashboard_html,
                        file_name="dashboard.html",
                        mime="text/html",
                        use_container_width=True,
                    )
                with dl2:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine="openpyxl") as writer:
                        pivot_summary.to_excel(
                            writer, sheet_name="Сводка по партиям"
                        )
                        pos_total = (
                            all_data.groupby(["category", "model"])["quantity"]
                            .sum()
                            .reset_index()
                            .rename(columns=COL_RU)
                        )
                        pos_total.sort_values(
                            by=["Категория", "Кол-во"],
                            ascending=[True, False],
                            inplace=True,
                        )
                        pos_total.to_excel(
                            writer, sheet_name="Всего по позициям", index=False
                        )
                        def_cols_x = [
                            "party", "category", "model", "serial_number",
                            "inv_number", "quantity", "condition", "engineer",
                            "date_updated",
                        ]
                        defective[def_cols_x].rename(columns=COL_RU).to_excel(
                            writer, sheet_name="Неудовлетворительные",
                            index=False,
                        )
                        violations.to_excel(
                            writer, sheet_name="Превышения лимитов",
                            index=False,
                        )
                        for ws in writer.book.worksheets:
                            for cell in ws[1]:
                                cell.font = Font(bold=True)
                            for col_cells in ws.columns:
                                max_len = max(
                                    (
                                        len(str(c.value))
                                        if c.value is not None
                                        else 0
                                    )
                                    for c in col_cells
                                )
                                ws.column_dimensions[
                                    get_column_letter(col_cells[0].column)
                                ].width = min(max_len + 2, 42)
                    output.seek(0)
                    st.download_button(
                        label="📥 Отчёт Excel (4 листа)",
                        data=output,
                        file_name="inventory_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                st.caption(
                    "HTML-дашборд открывается в браузере (графики, проценты, "
                    "блоки превышений и неисправной техники), его можно "
                    "распечатать в PDF: Ctrl+P → «Сохранить как PDF»."
                )

        with tab_all:
            st.markdown("### 📁 Реестр оргтехники")
            if all_data.empty:
                st.info("Реестр пуст.")
            else:
                f1, f2 = st.columns(2)
                with f1:
                    party_filter = st.multiselect(
                        "Фильтр по партиям",
                        options=sorted(all_data["party"].unique()),
                        placeholder="Все партии",
                    )
                with f2:
                    cat_filter = st.multiselect(
                        "Фильтр по категориям",
                        options=sorted(all_data["category"].unique()),
                        placeholder="Все категории",
                    )
                search = st.text_input(
                    "🔍 Поиск по реестру",
                    placeholder="Начните вводить: модель, серийный, "
                                "инвентарный номер, ответственный…",
                )

                filtered = all_data.copy()
                if party_filter:
                    filtered = filtered[filtered["party"].isin(party_filter)]
                if cat_filter:
                    filtered = filtered[filtered["category"].isin(cat_filter)]
                if search.strip():
                    q = search.strip().lower()
                    mask = (
                        filtered["model"].astype(str).str.lower()
                        .str.contains(q, regex=False, na=False)
                        | filtered["serial_number"].astype(str).str.lower()
                        .str.contains(q, regex=False, na=False)
                        | filtered["inv_number"].astype(str).str.lower()
                        .str.contains(q, regex=False, na=False)
                        | filtered["engineer"].astype(str).str.lower()
                        .str.contains(q, regex=False, na=False)
                        | filtered["category"].astype(str).str.lower()
                        .str.contains(q, regex=False, na=False)
                    )
                    filtered = filtered[mask]

                st.caption(
                    f"Показано позиций: {len(filtered)} из {len(all_data)}"
                )
                if filtered.empty:
                    st.warning("По заданным фильтрам ничего не найдено.")
                else:
                    st.dataframe(
                        style_conditions(
                            filtered.rename(columns=COL_RU)
                        ),
                        use_container_width=True,
                        hide_index=True,
                    )

                # Выгрузка общего списка в Excel
                output_all = io.BytesIO()
                with pd.ExcelWriter(output_all, engine="openpyxl") as writer:
                    all_data.to_excel(writer, index=False, sheet_name="Реестр")
                output_all.seek(0)
                st.download_button(
                    label="📥 Скачать полный реестр в Excel",
                    data=output_all,
                    file_name="all_equipment_registry.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                if not filtered.empty:
                    st.markdown("---")
                    st.markdown("### Удаление позиции из реестра")
                    del_options = {
                        f"ID {row.id} | Партия: {row.party} | {row.category} — "
                        f"{row.model} (Сер: {row.serial_number})": row.id
                        for _, row in filtered.iterrows()
                    }
                    selected_del_label = st.selectbox(
                        "Выберите позицию для удаления:",
                        list(del_options.keys()),
                    )
                    if st.button("🗑️ Удалить выбранную позицию"):
                        item_id_to_del = del_options[selected_del_label]
                        del_item = (
                            session.query(Equipment)
                            .filter(Equipment.id == item_id_to_del)
                            .first()
                        )
                        if del_item:
                            log_action(
                                session,
                                "Удаление",
                                del_item,
                                "Администратор",
                                "Позиция удалена администратором",
                            )
                        session.query(Equipment).filter(
                            Equipment.id == item_id_to_del
                        ).delete()
                        session.commit()
                        st.success("Позиция успешно удалена!")
                        st.rerun()

                if not filtered.empty:
                    st.markdown("---")
                    st.markdown("### 🏷 QR-коды для наклеек")
                    if not QR_AVAILABLE:
                        st.warning(
                            "Библиотека qrcode не установлена — QR-коды "
                            "недоступны. Добавьте пакет `qrcode` в "
                            "requirements.txt."
                        )
                    else:
                        base = app_base_url()
                        if not base:
                            st.info(
                                "Не удалось определить адрес приложения "
                                "для ссылки в QR-коде."
                            )
                        else:
                            qr_options = {
                                f"ID {row.id} | {row.party} | {row.category} — "
                                f"{row.model}": row.id
                                for _, row in filtered.iterrows()
                            }
                            qr_label = st.selectbox(
                                "Позиция для QR-кода",
                                list(qr_options.keys()),
                                key="qr_select",
                            )
                            qr_id = qr_options[qr_label]
                            qr_row = filtered[filtered.id == qr_id].iloc[0]
                            qr_url = f"{base}/?item={qr_id}"
                            qc1, qc2 = st.columns([1, 2])
                            with qc1:
                                st.markdown(
                                    f'<img src="{qr_data_uri(qr_url)}" '
                                    f'width="210" height="210" alt="QR-код">',
                                    unsafe_allow_html=True,
                                )
                            with qc2:
                                st.markdown(
                                    f"**{qr_row['category']} — "
                                    f"{qr_row['model']}**  \n"
                                    f"Партия: **{qr_row['party']}**  \n"
                                    f"Сер: {qr_row['serial_number']} · "
                                    f"Инв: {qr_row['inv_number']}"
                                )
                                st.code(qr_url, language=None)
                                st.caption(
                                    "Сканирование QR открывает карточку позиции."
                                )
                            dl_qr1, dl_qr2 = st.columns(2)
                            with dl_qr1:
                                st.download_button(
                                    "⬇️ Скачать QR (SVG)",
                                    data=qr_svg_bytes(qr_url),
                                    file_name=f"qr_item_{qr_id}.svg",
                                    mime="image/svg+xml",
                                    use_container_width=True,
                                )
                            with dl_qr2:
                                stickers = build_qr_stickers_html(
                                    filtered.to_dict("records"), base
                                )
                                st.download_button(
                                    "🖨 QR-наклейки для печати (HTML)",
                                    data=stickers,
                                    file_name="qr_stickers.html",
                                    mime="text/html",
                                    use_container_width=True,
                                )

                if IS_SQLITE and os.path.exists(DB_FILE):
                    st.markdown("---")
                    st.markdown("### 💾 Резервная копия базы")
                    st.caption(
                        "Скачивайте копию регулярно: при перезапуске приложения "
                        "на Streamlit Cloud временные файлы слота могут быть удалены. "
                        "Копию можно восстановить локально или перенести в Postgres "
                        "скриптом migrate_db.py."
                    )
                    with open(DB_FILE, "rb") as db_file:
                        st.download_button(
                            label="Скачать резервную копию базы (.db)",
                            data=db_file.read(),
                            file_name="inventory_backup.db",
                            mime="application/octet-stream",
                        )

        with tab_hist:
            st.markdown("### Журнал перемещений")
            if not history_data.empty:
                st.dataframe(
                    history_data.rename(columns=COL_RU),
                    use_container_width=True,
                    hide_index=True,
                )

                hist_out = io.BytesIO()
                with pd.ExcelWriter(hist_out, engine="openpyxl") as writer:
                    history_data.rename(columns=COL_RU).to_excel(
                        writer, sheet_name="История перемещений", index=False
                    )
                hist_out.seek(0)
                st.download_button(
                    label="📥 Скачать историю перемещений в Excel",
                    data=hist_out,
                    file_name="transfer_history.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.markdown("---")
                st.markdown("### Удаление записей истории")
                hist_del_options = {
                    f"ID {row.id} | Дата: {row.date} | {row.equipment_info} ({row.from_where} ➔ {row.to_where})": row.id
                    for _, row in history_data.iterrows()
                }
                selected_hist_label = st.selectbox(
                    "Выберите запись истории для удаления:",
                    list(hist_del_options.keys()),
                )
                if st.button("🗑️ Удалить выбранную запись истории"):
                    hist_id_to_del = hist_del_options[selected_hist_label]
                    session.query(History).filter(
                        History.id == hist_id_to_del
                    ).delete()
                    session.commit()
                    st.success("Запись истории успешно удалена!")
                    st.rerun()
            else:
                st.info("История перемещений пуста.")

            st.markdown("---")
            st.markdown(
                "### 🧾 Журнал действий (добавление / изменение / удаление)"
            )
            if not audit_data.empty:
                st.dataframe(
                    audit_data.rename(columns=COL_RU),
                    use_container_width=True,
                    hide_index=True,
                )
                audit_out = io.BytesIO()
                with pd.ExcelWriter(audit_out, engine="openpyxl") as writer:
                    audit_data.rename(columns=COL_RU).to_excel(
                        writer, sheet_name="Журнал действий", index=False
                    )
                audit_out.seek(0)
                st.download_button(
                    label="📥 Скачать журнал действий в Excel",
                    data=audit_out,
                    file_name="audit_log.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("Действий пока не зафиксировано.")

        with tab_move:
            st.markdown("### Перемещение техники")
            if not all_data.empty:
                move_search = st.text_input(
                    "🔍 Поиск позиции",
                    placeholder="Начните вводить: модель, серийный/инвентарный "
                                "номер, партия…",
                )
                move_pool = all_data
                if move_search.strip():
                    q = move_search.strip().lower()
                    mask = (
                        all_data["model"].astype(str).str.lower()
                        .str.contains(q, regex=False, na=False)
                        | all_data["serial_number"].astype(str).str.lower()
                        .str.contains(q, regex=False, na=False)
                        | all_data["inv_number"].astype(str).str.lower()
                        .str.contains(q, regex=False, na=False)
                        | all_data["category"].astype(str).str.lower()
                        .str.contains(q, regex=False, na=False)
                        | all_data["party"].astype(str).str.lower()
                        .str.contains(q, regex=False, na=False)
                    )
                    move_pool = all_data[mask]

                if move_pool.empty:
                    st.warning(
                        "По запросу ничего не найдено — уточните поиск."
                    )
                else:
                    item_options = {
                        f"ID {row.id} | {row.party} | {row.category} — "
                        f"{row.model} (Сер: {row.serial_number})": row.id
                        for _, row in move_pool.iterrows()
                    }
                    selected_item_label = st.selectbox(
                        "Выберите позицию", list(item_options.keys())
                    )
                    destination = st.selectbox(
                        "Куда переместить?", ["База"] + parties
                    )

                    st.markdown("#### Подтверждение ответственных")
                    st.caption(
                        "Оба поля обязательны: перемещение фиксируется на вас "
                        "и на принимающего."
                    )
                    admin_surname = st.text_input(
                        "Ваша фамилия (кто выполняет перемещение)",
                        placeholder="Например: Иванов",
                    )
                    receiver_surname = st.text_input(
                        "Фамилия принимающего (кому передаёте технику)",
                        placeholder="Например: Петров",
                    )

                    if st.button(
                        "Подтвердить перемещение", use_container_width=True
                    ):
                        ok, err = validate_transfer_surnames(
                            admin_surname,
                            receiver_surname,
                            current_engineer,
                            match_fio=False,
                        )
                        if not ok:
                            st.error("❌ " + err)
                        else:
                            item_id = item_options[selected_item_label]
                            item_to_move = (
                                session.query(Equipment)
                                .filter(Equipment.id == item_id)
                                .first()
                            )
                            if item_to_move:
                                if item_to_move.party == destination:
                                    st.error(
                                        "❌ Позиция уже находится в «"
                                        + destination
                                        + "». Выберите другое место."
                                    )
                                else:
                                    old_party = item_to_move.party
                                    item_to_move.party = destination
                                    item_to_move.date_updated = utc_now_str()

                                    history_entry = History(
                                        date=utc_now_str(),
                                        equipment_info=(
                                            f"{item_to_move.category} {item_to_move.model}"
                                        ),
                                        from_where=old_party,
                                        to_where=destination,
                                        engineer=admin_surname.strip(),
                                        receiver=receiver_surname.strip(),
                                        equipment_id=item_to_move.id,
                                    )
                                    session.add(history_entry)
                                    session.commit()
                                    st.success("Перемещение выполнено успешно!")
                                    st.rerun()
            else:
                st.info("В реестре нет техники для перемещения.")

    else:
        if password != "":
            st.error("Неверный пароль администратора!")
        st.warning("Введите пароль администратора в боковой панели слева.")

session.close()