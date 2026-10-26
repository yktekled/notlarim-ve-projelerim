#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
uretim_onay.py — ПРОИЗВОДСТВЕННОЕ СОГЛАСОВАНИЕ (üretim onay dosyası) PDF üreticisi.
YSF HICH / TEKLED. Sabit şablon: her çalıştırmada aynı düzen, elle düzen yok.

Kullanım:
    python3 uretim_onay.py proje.json cikti.pdf [--logo logo_tekled.png] [--rapor rapor.json]

Akış: JSON → kontrol motoru (STOP / KONTROL / INFO) → PDF (A4 yatay, RU).
STOP varsa kapakta "НЕ ЗАПУСКАТЬ В ПРОИЗВОДСТВО" otomatik basılır.
"""
import json, re, sys, os, argparse, io
from collections import Counter, defaultdict
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

# ───────────────────────── fontlar / renkler ─────────────────────────
FD = '/usr/share/fonts/truetype/dejavu/'
pdfmetrics.registerFont(TTFont('R', FD + 'DejaVuSans.ttf'))
pdfmetrics.registerFont(TTFont('B', FD + 'DejaVuSans-Bold.ttf'))
pdfmetrics.registerFont(TTFont('C', FD + 'DejaVuSansCondensed.ttf'))
pdfmetrics.registerFont(TTFont('CB', FD + 'DejaVuSansCondensed-Bold.ttf'))
from reportlab.pdfbase.pdfmetrics import registerFontFamily
registerFontFamily('C', normal='C', bold='CB', italic='C', boldItalic='CB')
registerFontFamily('R', normal='R', bold='B', italic='R', boldItalic='B')

RED = colors.HexColor('#C8102E'); AMBER = colors.HexColor('#B26B00'); GREEN = colors.HexColor('#2E7D32')
INK = colors.HexColor('#1F2329'); GREY = colors.HexColor('#6B7280'); LINE = colors.HexColor('#C9CED6')
HEAD = colors.HexColor('#2B2F36'); ZEBRA = colors.HexColor('#F3F5F8'); PALE_RED = colors.HexColor('#FDECEE')
PALE_AMB = colors.HexColor('#FFF4E0'); PANEL = colors.HexColor('#E9ECF1'); STEEL = colors.HexColor('#9AA3AE')

PW, PH = landscape(A4)
M = 18
FOOT_H = 46
TOP = PH - 78           # içerik üst sınırı
BOT = M + FOOT_H + 8    # içerik alt sınırı

def S(name, font='C', size=7.2, lead=None, color=INK, align=0):
    return ParagraphStyle(name, fontName=font, fontSize=size, leading=lead or size * 1.22, textColor=color, alignment=align)
P7 = S('p7'); P7B = S('p7b', 'CB'); P7W = S('p7w', 'CB', color=colors.white); P8 = S('p8', 'C', 8.2)
P8B = S('p8b', 'CB', 8.2); P6 = S('p6', 'C', 6.3); P7R = S('p7r', 'CB', color=RED); P7A = S('p7a', 'CB', color=AMBER)
P7C = S('p7c', align=1); P9 = S('p9', 'R', 9); P9B = S('p9b', 'B', 9)

def esc(s):
    return str('' if s is None else s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def para(s, st=P7):
    return Paragraph(esc(s) if not isinstance(s, Paragraph) else s, st) if not isinstance(s, Paragraph) else s

def fmt(v, suf=''):
    if v is None or v == '':
        return None
    if isinstance(v, float) and v == int(v):
        v = int(v)
    return f'{v}{suf}'

def yaz(c, x, y, txt, font, size, maxw, hiza='c', minsize=3.6):
    """Metni maxw genişliğine sığdırarak yazar: önce minsize'a kadar küçültür, yetmezse sonunu '…' ile keser.
    hiza: 'c' ortalı (x = merkez), 'l' sola (x = sol kenar)."""
    txt, s = str(txt or ''), size
    while s > minsize and pdfmetrics.stringWidth(txt, font, s) > maxw:
        s = round(s - 0.2, 2)
    if pdfmetrics.stringWidth(txt, font, s) > maxw:
        while txt and pdfmetrics.stringWidth(txt + '…', font, s) > maxw:
            txt = txt[:-1]
        txt = (txt + '…') if txt else ''
    c.setFont(font, s)
    (c.drawCentredString if hiza == 'c' else c.drawString)(x, y, txt)

KTL = 'КОНТРОЛЬ'
# TEKLED logosu (örnek PDF'lerden alındı) — dosya bağımsız çalışma için gömülü
LOGO_B64 = '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'

# ───────────────────────── sabit referanslar ─────────────────────────
KAT_RU = {'korpus': 'Корпус', 'rubilnik': 'Выключатель-разъединитель', 'mccb': 'Автомат. выключатель (литой корпус)',
          'mcb': 'Автомат. выключатель модульный', 'rccb': 'УЗО', 'rcbo': 'АВДТ', 'kontaktor': 'Контактор',
          'termik': 'Тепловое реле', 'opn': 'УЗИП', 'sigorta': 'Предохранитель', 'sayac': 'Счётчик эл. энергии',
          'tt': 'Трансформатор тока', 'lamba': 'Сигнальная лампа', 'buton': 'Кнопка', 'secici': 'Переключатель',
          'vfd': 'Преобразователь частоты / УПП', 'role': 'Реле / таймер', 'aksesuar': 'Аксессуар',
          'sarf': 'Монтажные изделия', 'montaj': 'Монтажный материал', 'iscilik': 'Монтажная работа', 'belirsiz': '—'}
DEVICE_CATS = {'mccb', 'mcb', 'rccb', 'rcbo', 'rubilnik', 'kontaktor', 'termik', 'opn', 'sigorta', 'sayac', 'tt',
               'lamba', 'buton', 'secici', 'vfd', 'role'}
DIN_CATS = {'mcb', 'rccb', 'rcbo', 'opn'}
# İç tesisat kesiti (Cu, PVC, pano içi) — XPAN pratiği; YALNIZCA genel.kesit_tablosu_onayli=true ise kullanılır
KESIT = [(25, 2.5), (32, 4), (40, 6), (50, 10), (63, 16), (80, 16), (100, 25), (125, 35), (160, 50), (200, 70), (250, 95)]
TORK = [('M4', 1.8, 1.5), ('M5', 5, 4.3), ('M6', 8, 6.8), ('M8', 20, 17), ('M10', 40, 35), ('M12', 70, 60), ('M16', 155, 130)]
ROUTINE = ['Степень защиты оболочки (IP) — визуально', 'Воздушные зазоры и пути утечки',
           'Защита от поражения током, непрерывность цепи PE', 'Встраивание комплектующих — по согласованной документации',
           'Внутренние цепи и соединения, момент затяжки', 'Выводы для внешних проводников',
           'Механическое функционирование (двери, замки, приводы, блокировки)',
           'Испытание электрической прочности изоляции', 'Проверка цепей управления и функционирования',
           'Маркировка: табличка НКУ, обозначения аппаратов, проводов, клемм']
# Kapı elemanı renk kuralları (IEC 60073 / IEC 60204-1 §10.2–10.3)
def renk_kontrol(e):
    y = (e.get('yazi') or '').lower(); r = (e.get('renk') or '').lower(); k = e.get('kategori')
    if not r:
        return ('KONTROL', f'{e["poz"]}: цвет не задан')
    if k == 'buton':
        if re.search(r'авар|аварийн|e-?stop|emergency', y) and 'красн' not in r:
            return ('STOP', f'{e["poz"]} «{e.get("yazi")}»: аварийный стоп — только красный (на жёлтом фоне)')
        if re.search(r'стоп|stop|откл', y) and 'зел' in r:
            return ('STOP', f'{e["poz"]} «{e.get("yazi")}»: кнопка СТОП не может быть зелёной')
        if re.search(r'пуск|start|вкл', y) and 'красн' in r:
            return ('STOP', f'{e["poz"]} «{e.get("yazi")}»: кнопка ПУСК не может быть красной')
    if k == 'lamba':
        if re.search(r'авар|fault|alarm|неиспр', y) and 'красн' not in r:
            return ('KONTROL', f'{e["poz"]} «{e.get("yazi")}»: лампа «Авария» по IEC 60073 — красная')
    return None


# ───────────────────────── kontrol motoru ─────────────────────────
def modul(d):
    if d.get('modul'):
        return float(d['modul'])
    if d.get('kategori') not in DIN_CATS:
        return None
    k = (d.get('kutup') or '').upper()
    m = re.match(r'(\d)P(\+N)?', k)
    if not m:
        return None
    n = int(m.group(1)) + (1 if m.group(2) else 0)
    return float(n)

def kesit_mm2(kablo):
    m = re.search(r'\d\s*[xх×]\s*(\d+(?:[.,]\d+)?)', kablo or '')
    return float(m.group(1).replace(',', '.')) if m else None

def kontrol_motoru(data):
    """Her pano için sorun listesi döndürür: (seviye, kod, mesaj, poz)."""
    genel = data.get('genel', {})
    proje_sorun = []
    if not genel.get('bara_renk'):
        proje_sorun.append(('KONTROL', 'G01', 'Цветовая маркировка шин не выбрана: ПУЭ п.1.1.30 (L1 жёлтый / L2 зелёный / L3 красный) '
                            'или IEC 60445 (коричневый / чёрный / серый). Подтвердить заказчиком.', ''))
    if not genel.get('ic_faz_renk'):
        proje_sorun.append(('KONTROL', 'G03', 'Цвет фазных проводов внутренней разводки не выбран: IEC 60445 (L1 коричневый / L2 чёрный / L3 серый) '
                            'или все фазы чёрные.', ''))
    if not genel.get('kumanda_renk') and any(p.get('kapi_elemanlari') or any(d.get('kategori') in ('kontaktor', 'role', 'vfd') for d in p.get('devreler', []))
                                              for p in data['panolar']):
        proje_sorun.append(('KONTROL', 'G04', 'Цвета проводов цепей управления не заданы (~230 В L/N, =24 В +/−).', ''))
    if not data['proje'].get('kaynak_sema'):
        proje_sorun.append(('KONTROL', 'G02', 'Однолинейная схема не приложена — названия линий и кабели взяты только из BOM.', ''))
    sonuc = {}
    for p in data['panolar']:
        s = []
        t = p.get('teknik', {}); k = p.get('korpus', {})
        zorunlu = [('H', k.get('H'), 'Высота корпуса'), ('W', k.get('W'), 'Ширина корпуса'), ('D', k.get('D'), 'Глубина корпуса'),
                   ('IP', k.get('IP'), 'Степень защиты IP'), ('Ue', t.get('Ue'), 'Номинальное напряжение Ue'),
                   ('sistem', t.get('sistem'), 'Система заземления'), ('In_pano', t.get('In_pano'), 'Номинальный ток НКУ InA')]
        eksik = [ad for _, v, ad in zorunlu if v in (None, '')]
        if eksik:
            s.append(('STOP', 'R01', 'Не заданы обязательные данные: ' + ', '.join(eksik), ''))
        istege = [(t.get('Ik_proje_kA'), 'ток КЗ в точке установки Ik'), (k.get('montaj'), 'способ установки'),
                  (k.get('RAL'), 'цвет RAL'), (k.get('kablo_giris'), 'кабельный ввод'), (t.get('form'), 'форма секционирования'),
                  (t.get('klimatik'), 'климатическое исполнение'), (t.get('Ui'), 'Ui')]
        eksik2 = [ad for v, ad in istege if v in (None, '')]
        if eksik2:
            s.append(('KONTROL', 'R02', 'Уточнить: ' + ', '.join(eksik2), ''))
        devs = p.get('devreler', []); kapi = p.get('kapi_elemanlari', [])
        # R10 tekrar eden poz
        for poz, n in Counter(d['poz'] for d in devs + kapi if d.get('poz')).items():
            if n > 1:
                s.append(('STOP', 'R10', f'Обозначение {poz} повторяется {n} раза', poz))
        # R03/R04 BOM ↔ tablo
        bom = defaultdict(float); bom_ad = {}
        for m in p.get('malzemeler', []):
            if m.get('kategori') in DEVICE_CATS and m.get('artikul'):
                bom[m['artikul']] += m.get('adet') or 0; bom_ad[m['artikul']] = m.get('ad')
        tab = Counter(d.get('artikul') for d in devs + kapi if d.get('artikul'))
        for art in sorted(set(bom) | set(tab)):
            b, tq = bom.get(art, 0), tab.get(art, 0)
            if art not in bom:
                s.append(('STOP', 'R04', f'Арт. {art}: есть в таблице линий ({tq} шт), нет в BOM', ''))
            elif abs(b - tq) > 1e-6:
                s.append(('STOP', 'R03', f'Арт. {art} «{bom_ad.get(art, "")}»: BOM {fmt(b)} шт ≠ таблица линий {tq} шт', ''))
        # R06 Icu
        Ik = t.get('Ik_proje_kA')
        for d in devs:
            if Ik and d.get('Icu_kA') and d.get('grup') == 'giris' and float(d['Icu_kA']) < float(Ik):
                s.append(('STOP', 'R06', f'{d["poz"]}: Icu {fmt(d["Icu_kA"])} кА < Ik {fmt(Ik)} кА', d['poz']))
        # R07 çıkış > giriş
        giris = [d for d in devs if d.get('grup') == 'giris' and d.get('In')]
        if giris:
            gmax = max(float(d['In']) for d in giris)
            for d in devs:
                if d.get('grup') == 'cikis' and d.get('In') and float(d['In']) > gmax:
                    s.append(('KONTROL', 'R07', f'{d["poz"]}: In {fmt(d["In"])} А больше вводного {fmt(gmax)} А (селективность)', d['poz']))
        elif devs:
            s.append(('KONTROL', 'R07', 'Вводной аппарат не отмечен (grup = giris)', ''))
        # R08 kesit
        if genel.get('kesit_tablosu_onayli'):
            for d in devs:
                q = kesit_mm2(d.get('ic_kablo')); In = d.get('In')   # yalnız pano içi iletken; saha kablosu proje kararıdır
                if q and In:
                    need = next((x for lim, x in KESIT if float(In) <= lim), None)
                    if need and q < need:
                        s.append(('KONTROL', 'R08', f'{d["poz"]}: внутренний провод {q} мм² < {need} мм² для {fmt(In)} А', d['poz']))
        # R09 renkler
        for e in kapi:
            r = renk_kontrol(e)
            if r:
                s.append((r[0], 'R09', r[1], e['poz']))
        # R11 RCD
        for d in devs:
            if d.get('kategori') in ('rccb', 'rcbo') and (not d.get('rcd_mA') or not d.get('rcd_tip')):
                s.append(('KONTROL', 'R11', f'{d["poz"]}: не указан IΔn или тип УЗО (AC/A/B)', d['poz']))
        # R12 yük adı
        bos = [d['poz'] for d in devs if d.get('grup') == 'cikis' and not d.get('yuk') and d.get('kategori') in ('mcb', 'mccb', 'rcbo', 'rubilnik')]
        if bos:
            s.append(('KONTROL', 'R12', 'Нет наименования линии: ' + ', '.join(bos[:20]) + (' …' if len(bos) > 20 else ''), ''))
        # R05 modül doluluk
        if k.get('din_sira') and k.get('sira_modul'):
            cap = int(k['din_sira']) * float(k['sira_modul'])
            used = sum(modul(d) or 0 for d in devs) + float(k.get('ek_modul') or 0)
            if used > cap:
                s.append(('STOP', 'R05', f'Модульные аппараты {fmt(used)} мод. > ёмкость {fmt(cap)} мод. Корпус мал.', ''))
            elif cap and (cap - used) / cap < 0.20:
                s.append(('KONTROL', 'R05', f'Резерв {100 * (cap - used) / cap:.0f}% < 20% ({fmt(used)}/{fmt(cap)} мод.)', ''))
        elif any(modul(d) for d in devs):
            s.append(('KONTROL', 'R05', 'Не задано кол-во DIN-рядов / модулей корпуса — заполнение не проверено', ''))
        nd = [d['poz'] for d in devs if d.get('kategori') in DIN_CATS and modul(d) is None]
        if nd:
            s.append(('KONTROL', 'R05', 'Ширина в модулях не определена: ' + ', '.join(nd[:15]), ''))
        # elle girilen
        for kk in p.get('kontrol', []):
            s.append((kk.get('seviye', 'KONTROL'), 'M', kk['mesaj'], kk.get('poz', '')))
        sonuc[p['kod']] = s
    return proje_sorun, sonuc


# ───────────────────────── çizim yardımcıları ─────────────────────────
class Doc:
    def __init__(self, data, logo, total=None):
        self.d = data; self.logo = logo; self.total = total; self.n = 0
        self.buf = io.BytesIO()
        self.c = rl_canvas.Canvas(self.buf, pagesize=(PW, PH))
        pr = data['proje']
        self.c.setTitle(f'{pr.get("no")} Производственное согласование {pr.get("rev")}')
        self.c.setAuthor(pr.get('firma', 'TEKLED'))

    def page(self, title, sub, section, pano=None, status=None):
        c = self.c; self.n += 1; pr = self.d['proje']
        src = self.logo if (self.logo and os.path.exists(self.logo)) else (io.BytesIO(__import__('base64').b64decode(LOGO_B64)) if LOGO_B64 else None)
        if src:
            c.drawImage(ImageReader(src), M, PH - M - 26, width=97, height=26, mask='auto')
        else:
            c.setFont('B', 16); c.setFillColor(RED); c.drawString(M, PH - M - 20, pr.get('firma', 'TEKLED'))
        c.setFillColor(INK); c.setFont('B', 12.5); c.drawRightString(PW - M, PH - M - 12, title)
        c.setFont('C', 7.5); c.setFillColor(GREY); c.drawRightString(PW - M, PH - M - 24, sub)
        c.setStrokeColor(RED); c.setLineWidth(1.2); c.line(M, PH - M - 33, PW - M, PH - M - 33)
        c.setFillColor(INK); c.setFont('B', 10.5); c.drawString(M, PH - M - 50, section)
        st = status or pr.get('durum', 'НА СОГЛАСОВАНИЕ')
        c.setFont('CB', 8.5); c.setFillColor(RED if 'НЕ ' in st else INK)
        c.drawRightString(PW - M, PH - M - 50, f'{pr.get("rev")} | {pr.get("tarih")} | {st}')
        self.footer(section, pano)

    def footer(self, section, pano):
        c = self.c; pr = self.d['proje']; y0 = M; h = FOOT_H
        cols = [('Разработал', pr.get('hazirlayan', '')), ('Проверил', pr.get('kontrol_eden', '')), ('Производство', ''),
                ('ОТК', ''), ('Согласовано заказчиком', '')]
        w = [88, 88, 88, 88, 150]; x = M
        c.setLineWidth(0.5); c.setStrokeColor(GREY)
        for (lab, val), ww in zip(cols, w):
            c.rect(x, y0, ww, h)
            c.setFont('C', 6.3); c.setFillColor(GREY); c.drawString(x + 3, y0 + h - 9, lab)
            c.setFont('CB', 7); c.setFillColor(INK); c.drawString(x + 3, y0 + h - 20, val)
            c.setFont('C', 5.6); c.setFillColor(GREY); c.drawString(x + 3, y0 + 4, 'Ф.И.О. / подпись / дата')
            x += ww
        wmid = PW - M - x - 150
        c.rect(x, y0, wmid, h)
        c.setFillColor(INK)
        yaz(c, x + 5, y0 + h - 12, (pano['kod'] + ' | ' + pano['ad']) if pano else f'Проект {pr.get("no")}', 'CB', 8, wmid - 10, 'l', 6)
        c.setFillColor(GREY)
        yaz(c, x + 5, y0 + h - 23, section, 'C', 6.6, wmid - 10, 'l', 5.6)
        yaz(c, x + 5, y0 + 5, f'Источник: {pr.get("kaynak_bom", "")}', 'C', 6.6, wmid - 10, 'l', 5.6)
        x += wmid
        cells = [('Проект', pr.get('no', '')), ('Лист', f'{self.n} / {self.total or "?"}'), ('Рев.', pr.get('rev', '')), ('Дата', pr.get('tarih', ''))]
        cw = 150 / 4
        for i, (lab, val) in enumerate(cells):
            c.rect(x + i * cw, y0, cw, h)
            c.setFont('C', 6); c.setFillColor(GREY); c.drawCentredString(x + i * cw + cw / 2, y0 + h - 10, lab)
            c.setFillColor(INK)
            yaz(c, x + i * cw + cw / 2, y0 + 16, str(val), 'CB', 7.2 if lab != 'Лист' else 7.6, cw - 4, 'c', 5.5)
        c.setFillColor(INK)

    def end(self):
        self.c.showPage()

    def table(self, rows, widths, x, ytop, style_extra=None, head=1, zebra=True):
        t = Table(rows, colWidths=widths, repeatRows=head)
        st = [('FONT', (0, 0), (-1, -1), 'C', 7), ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
              ('GRID', (0, 0), (-1, -1), 0.4, LINE), ('TOPPADDING', (0, 0), (-1, -1), 2.2),
              ('BOTTOMPADDING', (0, 0), (-1, -1), 2.2), ('LEFTPADDING', (0, 0), (-1, -1), 3), ('RIGHTPADDING', (0, 0), (-1, -1), 3)]
        if head:
            st += [('BACKGROUND', (0, 0), (-1, head - 1), HEAD)]
        if zebra:
            for i in range(head, len(rows)):
                if (i - head) % 2:
                    st.append(('BACKGROUND', (0, i), (-1, i), ZEBRA))
        t.setStyle(TableStyle(st + (style_extra or [])))
        _, h = t.wrap(sum(widths), 10000)
        t.drawOn(self.c, x, ytop - h)
        return h

    def note_box(self, x, y, w, title, lines, col=RED, bg=PALE_RED):
        ps = [Paragraph(f'<font name="CB" color="#{col.hexval()[2:]}">{esc(title)}</font>', P7)]
        ps += [Paragraph(esc(l), P7) for l in lines]
        hs = [pp.wrap(w - 12, 1000)[1] + 2 for pp in ps]; h = sum(hs) + 10
        c = self.c; c.setFillColor(bg); c.setStrokeColor(col); c.setLineWidth(0.8)
        c.roundRect(x, y - h, w, h, 4, fill=1, stroke=1)
        yy = y - 6
        for pp, hh in zip(ps, hs):
            pp.drawOn(c, x + 6, yy - hh + 2); yy -= hh
        c.setFillColor(INK)
        return h


def chunk_rows(header, rows, widths, avail, extra_style_fn=None):
    """Satırları sayfalara böler (ölçerek)."""
    out, cur = [], []
    hh = Table([header], colWidths=widths).wrap(sum(widths), 1000)[1]
    used = hh
    for r in rows:
        rh = Table([r], colWidths=widths, style=[('FONT', (0, 0), (-1, -1), 'C', 7), ('TOPPADDING', (0, 0), (-1, -1), 2.2),
                                                 ('BOTTOMPADDING', (0, 0), (-1, -1), 2.2)]).wrap(sum(widths), 1000)[1]
        if cur and used + rh > avail:
            out.append(cur); cur = []; used = hh
        cur.append(r); used += rh
    if cur or not out:
        out.append(cur)
    return out


def hdr(cells):
    return [Paragraph(esc(x), P7W) for x in cells]


def val_or_ktl(v, suf=''):
    f = fmt(v, suf)
    return Paragraph(esc(f), P7) if f is not None else Paragraph(KTL, P7R)


# ───────────────────────── sayfalar ─────────────────────────
def durum(sorunlar):
    if any(s[0] == 'STOP' for s in sorunlar):
        return 'НЕ ЗАПУСКАТЬ В ПРОИЗВОДСТВО ДО ПОДТВЕРЖДЕНИЯ'
    if any(s[0] == 'KONTROL' for s in sorunlar):
        return 'НА СОГЛАСОВАНИЕ (ЕСТЬ ВОПРОСЫ)'
    return 'НА СОГЛАСОВАНИЕ'


def sub_line(d, pano=None):
    pr = d['proje']
    parts = [f'Проект {pr.get("no")}', pr.get('sehir'), pr.get('musteri')]
    if pano:
        parts += [pano['kod'], pano['ad']]
    return ' | '.join(str(x) for x in parts if x)


def kapak(doc, proje_sorun, sonuc, sayfa_map):
    d = doc.d; pr = d['proje']; c = doc.c
    tum = proje_sorun + [x for v in sonuc.values() for x in v]
    st = durum(tum)
    doc.page('ПРОИЗВОДСТВЕННОЕ СОГЛАСОВАНИЕ', sub_line(d), 'КОМПЛЕКТ ЩИТОВ — ТИТУЛЬНЫЙ ЛИСТ', status=st)
    y = TOP
    c.setFont('B', 17); c.setFillColor(INK); c.drawString(M, y - 8, (pr.get('musteri') or '').upper())
    c.setFont('R', 9); c.setFillColor(GREY)
    c.drawString(M, y - 24, ' | '.join(x for x in [pr.get('no'), pr.get('sehir'), pr.get('obje')] if x))
    rows = [hdr(['№', 'Код', 'Щит / наименование', 'Габариты В×Ш×Г, мм', 'IP', 'Ввод', 'Кол.', 'Листы', 'Стоп', 'Контр.'])]
    for i, p in enumerate(d['panolar'], 1):
        k = p.get('korpus', {}); s = sonuc[p['kod']]
        gab = '×'.join(str(k.get(x)) for x in 'HWD') if all(k.get(x) for x in 'HWD') else None
        gi = next((dd for dd in p.get('devreler', []) if dd.get('grup') == 'giris'), None)
        ns = sum(1 for x in s if x[0] == 'STOP'); nk = sum(1 for x in s if x[0] == 'KONTROL')
        rows.append([para(i), para(p['kod']), para(f'{p["ad"]}' + (f' — {p["aciklama"]}' if p.get('aciklama') else '')),
                     val_or_ktl(gab), val_or_ktl(k.get('IP')),
                     para(f'{fmt(gi.get("In"))} А {gi.get("kutup") or ""}' if gi and gi.get('In') else '—'),
                     para(f'{p.get("adet", 1)} компл.'), para(sayfa_map.get(p['kod'], '')),
                     Paragraph(str(ns), P7R if ns else P7), Paragraph(str(nk), P7A if nk else P7)])
    widths = [22, 38, 250, 105, 40, 70, 52, 56, 40, 40]
    pages = chunk_rows(rows[0], rows[1:], widths, y - 42 - BOT - 120)
    # durum bandı
    n_stop = sum(1 for s in tum if s[0] == 'STOP'); n_k = sum(1 for s in tum if s[0] == 'KONTROL')
    bg, col = (PALE_RED, RED) if n_stop else ((PALE_AMB, AMBER) if n_k else (colors.HexColor('#E8F5E9'), GREEN))
    c.setFillColor(bg); c.setStrokeColor(col); c.roundRect(PW - M - 330, y - 30, 330, 32, 4, fill=1, stroke=1)
    c.setFillColor(col); c.setFont('CB', 9); c.drawString(PW - M - 322, y - 12, st)
    c.setFont('C', 7.5); c.drawString(PW - M - 322, y - 24, f'Стоп-пунктов: {n_stop}   |   Вопросов КОНТРОЛЬ: {n_k}   |   Реестр — лист {len(pages) + 1}')
    y -= 42
    h = doc.table([rows[0]] + pages[0], widths, M, y)
    y -= h + 10
    if len(pages) > 1:
        c.setFont('C', 7); c.setFillColor(AMBER); c.drawString(M, y, f'Продолжение перечня — на следующем листе ({len(rows) - 1 - len(pages[0])} щитов).'); y -= 12
    # revizyon geçmişi + kaynak
    rg = pr.get('rev_gecmisi', [])
    rrows = [hdr(['Рев.', 'Дата', 'Изменение'])] + [[para(r.get('rev')), para(r.get('tarih')), para(r.get('aciklama'))] for r in rg[-5:]]
    hr = doc.table(rrows, [36, 60, 330], M, y)
    notes = [f'Исходные данные: BOM «{pr.get("kaynak_bom", "")}»' + (f'; схема «{pr["kaynak_sema"]}»' if pr.get('kaynak_sema') else '; однолинейная схема — не приложена'),
             'Документ содержит состав без цен. Количество аппаратов сверено с BOM автоматически (артикул × кол-во).',
             'Пометка КОНТРОЛЬ — данные требуют подтверждения. Стоп-пункт — изготовление не начинать до решения заказчика.',
             'Изготовление после подписи заказчика на каждом титульном листе щита и в реестре вопросов.']
    doc.note_box(M + 440, y, PW - 2 * M - 440, 'ПРИМЕЧАНИЯ', notes, GREY, ZEBRA)
    doc.end()
    for extra in pages[1:]:
        doc.page('ПРОИЗВОДСТВЕННОЕ СОГЛАСОВАНИЕ', sub_line(d), 'КОМПЛЕКТ ЩИТОВ — ПЕРЕЧЕНЬ (ПРОДОЛЖЕНИЕ)', status=st)
        doc.table([rows[0]] + extra, widths, M, TOP); doc.end()


def register(doc, proje_sorun, sonuc):
    d = doc.d
    items = [('Проект', *s) for s in proje_sorun]
    for p in d['panolar']:
        items += [(f'{p["kod"]} {p["ad"]}', *s) for s in sorted(sonuc[p['kod']], key=lambda x: (x[0] != 'STOP', x[1]))]
    rows_h = hdr(['№', 'Щит', 'Уровень', 'Код', 'Вопрос / расхождение', 'Решение заказчика'])
    rows = []
    for i, (pn, lev, kod, msg, poz) in enumerate(items, 1):
        rows.append([para(f'K-{i:02d}'), para(pn), Paragraph('СТОП' if lev == 'STOP' else ('КОНТРОЛЬ' if lev == 'KONTROL' else 'ИНФО'),
                     P7R if lev == 'STOP' else P7A), para(kod), para(msg), para('')])
    widths = [34, 120, 52, 30, 420, PW - 2 * M - 656]
    if not rows:
        doc.page('РЕЕСТР ВОПРОСОВ', sub_line(d), 'РЕЕСТР ВОПРОСОВ И СТОП-ПУНКТОВ')
        doc.c.setFont('CB', 11); doc.c.setFillColor(GREEN); doc.c.drawString(M, TOP - 20, 'Открытых вопросов нет.'); doc.end(); return
    for j, ch in enumerate(chunk_rows(rows_h, rows, widths, TOP - BOT - 30)):
        doc.page('РЕЕСТР ВОПРОСОВ', sub_line(d), 'РЕЕСТР ВОПРОСОВ И СТОП-ПУНКТОВ' + (' (продолжение)' if j else ''))
        h = doc.table([rows_h] + ch, widths, M, TOP)
        doc.c.setFont('C', 7); doc.c.setFillColor(GREY)
        doc.c.drawString(M, TOP - h - 12, 'Ответ заказчика вносится в колонку «Решение». После ответа выпускается следующая ревизия с отметкой о закрытии вопроса.')
        doc.end()


def genel_gereksinim(doc):
    d = doc.d; g = d.get('genel', {}); c = doc.c
    doc.page('ОБЩИЕ ТЕХНИЧЕСКИЕ ТРЕБОВАНИЯ', sub_line(d), 'ОБЩИЕ ТРЕБОВАНИЯ К ИЗГОТОВЛЕНИЮ И МАРКИРОВКЕ')
    colw = (PW - 2 * M - 20) / 2; xL = M; xR = M + colw + 20; y = TOP
    # renkler
    br = g.get('bara_renk')
    bara = {'ПУЭ': ('жёлтый', 'зелёный', 'красный'), 'IEC': ('коричневый', 'чёрный', 'серый')}.get(br)
    rows = [hdr(['Проводник', 'Провода (ГОСТ 33542 / IEC 60445)', 'Шины'])]
    ifr = {'IEC': 'коричневый / чёрный / серый', 'BK': 'чёрный (все фазы)'}.get(g.get('ic_faz_renk'))
    kr = g.get('kumanda_renk')
    rows += [[para('Фаза L1 / L2 / L3'), para(ifr) if ifr else Paragraph(KTL + ': выбрать', P7R),
              para(' / '.join(bara)) if bara else Paragraph(KTL + ': выбрать ПУЭ или IEC', P7R)],
             [para('Нейтраль N'), para('синий (голубой)'), para('синий (голубой)')],
             [para('Защитный PE'), para('жёлто-зелёный'), para('жёлто-зелёные полосы')],
             [para('PEN'), para('жёлто-зелёный + голубые метки на концах'), para('голубой + жёлто-зелёные полосы')],
             [para('Цепи управления ~230 В / =24 В'), para(kr) if kr else Paragraph(KTL + ': по схеме проекта / задать', P7A), para('—')]]
    h = doc.table(rows, [110, 170, colw - 280], xL, y); y -= h + 8
    rows = [hdr(['Обозн.', 'Аппарат (ГОСТ 2.710)', 'Обозн.', 'Аппарат'])]
    pairs = [('QS', 'Выключатель-разъединитель'), ('QF', 'Автоматический выключатель'), ('QFD', 'АВДТ'), ('QD', 'УЗО'),
             ('KM', 'Контактор'), ('KK', 'Тепловое реле'), ('KL/KT', 'Реле / реле времени'), ('FV', 'УЗИП'),
             ('FU', 'Предохранитель'), ('PI', 'Счётчик'), ('TA', 'Трансформатор тока'), ('HL', 'Сигнальная лампа'),
             ('SB', 'Кнопка'), ('SA', 'Переключатель'), ('UZ', 'Преобразователь частоты'), ('XT', 'Клеммник')]
    for i in range(0, len(pairs), 2):
        a, b = pairs[i], pairs[i + 1]
        rows.append([para(a[0], P7B), para(a[1]), para(b[0], P7B), para(b[1])])
    h = doc.table(rows, [40, colw / 2 - 40, 40, colw / 2 - 40], xL, y); y -= h + 8
    if g.get('kesit_tablosu_onayli'):
        rows = [hdr(['In, А'] + [f'≤{a}' for a, _ in KESIT]), [para('Cu, мм²', P7B)] + [para(fmt(q)) for _, q in KESIT]]
        c.setFont('CB', 7.5); c.setFillColor(INK); c.drawString(xL, y - 2, 'Сечение внутренней разводки (Cu, ПВ3 / H07V-K) по току аппарата'); y -= 8
        w0 = 44; doc.table(rows, [w0] + [(colw - w0) / len(KESIT)] * len(KESIT), xL, y)
    # sağ: etiketleme + rutin testler
    y = TOP
    lab = ['Табличка НКУ (ГОСТ IEC 61439-1 п.6.1): изготовитель / товарный знак; обозначение типа или заводской номер; '
           'дата изготовления; обозначение стандарта ГОСТ IEC 61439-X; знак EAC (ТР ТС 004/2011, 020/2011).',
           'На табличке или в паспорте: Ue, InA, Icw / Ipk, частота, IP, система заземления, масса при необходимости.',
           'Каждый аппарат — обозначение по таблице линий (напр. QF3) и наименование линии на лицевой панели / под рукояткой.',
           'Провода — маркировка с обеих сторон (номер цепи / клемма аппарата). Клеммы — обозначение XT и номер.',
           'Знак «Осторожно! Электрическое напряжение» на двери (ГОСТ 12.4.026). Ввод питания — табличка «ВВОД».',
           'Однолинейная схема и перечень линий — в кармане двери (комплект документации).']
    hb = doc.note_box(xR, y, colw, 'МАРКИРОВКА', lab, GREY, ZEBRA); y -= hb + 8
    rows = [hdr(['№', 'Приёмо-сдаточные проверки (ГОСТ IEC 61439-1 р.11)', 'Отметка ОТК'])]
    rows += [[para(i), para(r), para('☐')] for i, r in enumerate(ROUTINE, 1)]
    h = doc.table(rows, [20, colw - 80, 60], xR, y); y -= h + 8
    if g.get('tork_tablosu_onayli'):
        rows = [hdr(['Болт (кл. 8.8)', 'Затяжка, Н·м', 'Контроль, Н·м'])] + [[para(a), para(fmt(b)), para(fmt(cc))] for a, b, cc in TORK]
        doc.table(rows, [80, 80, 80], xR, y)
    doc.end()


def front_sketch(c, x, y, w, h, k, kapi, title=None, open_=False):
    """Orantılı ön görünüş. (x,y)=sol-alt, w/h=kutu. Döndürür ölçek."""
    H, W = k.get('H') or 600, k.get('W') or 400
    sc = min(w / W, h / H); ww, hh = W * sc, H * sc; x0 = x + (w - ww) / 2; y0 = y + (h - hh) / 2
    c.setStrokeColor(INK); c.setLineWidth(1); c.setFillColor(PANEL); c.rect(x0, y0, ww, hh, fill=1)
    c.setFillColor(colors.white); c.rect(x0 + 5, y0 + 5, ww - 10, hh - 10, fill=1)
    if not open_:
        c.setFillColor(STEEL)
        c.rect(x0 + ww - 3, y0 + hh * 0.2, 4, 10, fill=1); c.rect(x0 + ww - 3, y0 + hh * 0.75, 4, 10, fill=1)  # menteşe sağ
        c.setFillColor(colors.HexColor('#F2B705')); c.circle(x0 + 9, y0 + hh / 2, 3, fill=1)                  # kilit sol
        # kapı elemanları
        n = len(kapi)
        if n:
            per = max(1, min(n, int((ww - 20) // 26)))
            for i, e in enumerate(kapi):
                r_, col_ = divmod(i, per)
                cx = x0 + ww / 2 - (min(per, n) - 1) * 13 + col_ * 26; cy = y0 + hh * 0.82 - r_ * 26
                rn = (e.get('renk') or '').lower()
                fc = {'красн': '#D32F2F', 'зел': '#2E7D32', 'жёлт': '#F2B705', 'желт': '#F2B705', 'бел': '#FFFFFF',
                      'син': '#1565C0', 'чёрн': '#111111', 'черн': '#111111'}
                col = next((v for kk, v in fc.items() if kk in rn), '#BBBBBB')
                c.setFillColor(colors.HexColor(col)); c.setStrokeColor(INK)
                if e.get('kategori') == 'secici':
                    c.rect(cx - 5, cy - 5, 10, 10, fill=1)
                else:
                    c.circle(cx, cy, 5, fill=1)
                c.setFillColor(INK); yaz(c, cx, cy - 11, e.get('yazi') or e['poz'], 'C', 4.6, 25)
        # uyarı işareti
        c.setFillColor(colors.HexColor('#F2B705')); tx, ty = x0 + ww / 2, y0 + hh * 0.55
        c.setStrokeColor(INK); p = c.beginPath(); p.moveTo(tx - 7, ty - 6); p.lineTo(tx + 7, ty - 6); p.lineTo(tx, ty + 7); p.close()
        c.drawPath(p, fill=1, stroke=1); c.setFillColor(INK); c.setFont('B', 6); c.drawCentredString(tx, ty - 4, '⚡')
    c.setFillColor(INK)
    return sc, x0, y0, ww, hh


def dim_h(c, x1, x2, y, txt):
    c.setStrokeColor(GREY); c.setLineWidth(0.4); c.line(x1, y, x2, y); c.line(x1, y - 3, x1, y + 3); c.line(x2, y - 3, x2, y + 3)
    c.setFont('C', 6.5); c.setFillColor(INK); c.drawCentredString((x1 + x2) / 2, y - 9, txt)


def dim_v(c, x, y1, y2, txt):
    c.setStrokeColor(GREY); c.setLineWidth(0.4); c.line(x, y1, x, y2); c.line(x - 3, y1, x + 3, y1); c.line(x - 3, y2, x + 3, y2)
    c.saveState(); c.translate(x - 4, (y1 + y2) / 2); c.rotate(90); c.setFont('C', 6.5); c.setFillColor(INK)
    c.drawCentredString(0, 0, txt); c.restoreState()


def pano_titul(doc, p, sorun):
    c = doc.c; k = p.get('korpus', {}); t = p.get('teknik', {})
    st = durum(sorun)
    doc.page('ПРОИЗВОДСТВЕННОЕ СОГЛАСОВАНИЕ', sub_line(doc.d, p), 'ТИТУЛЬНЫЙ ЛИСТ ЩИТА И ТЕХНИЧЕСКИЕ ДАННЫЕ', p, st)
    y = TOP
    c.setFont('B', 15); c.setFillColor(INK); c.drawString(M, y - 10, p['ad'])
    c.setFont('C', 8); c.setFillColor(GREY); c.drawString(M, y - 23, (p.get('aciklama') or '')[:110])
    front_sketch(c, M, BOT + 70, 230, y - 40 - BOT - 80, k, p.get('kapi_elemanlari', []))
    c.setFont('C', 7); c.setFillColor(GREY); c.drawCentredString(M + 115, BOT + 60, 'ФАСАД — ДВЕРЬ ЗАКРЫТА (пропорционально)')
    gi = next((d for d in p.get('devreler', []) if d.get('grup') == 'giris'), None)
    gab = f'{k["H"]} × {k["W"]} × {k["D"]} мм' if all(k.get(x) for x in 'HWD') else None
    par = [('Наименование щита', f'{p["ad"]}'), ('Количество', f'{p.get("adet", 1)} компл.'),
           ('Корпус (тип, артикул, изготовитель)', ' '.join(x for x in [k.get('ad'), k.get('artikul') and f'арт. {k["artikul"]}', k.get('marka')] if x) or None),
           ('Габариты В × Ш × Г', gab), ('Установка', k.get('montaj')), ('Степень защиты IP / IK', ' / '.join(x for x in [k.get('IP'), t.get('IK')] if x) or None),
           ('Цвет', f'RAL {k["RAL"]}' if k.get('RAL') else None), ('Климатическое исполнение / место', ' / '.join(x for x in [t.get('klimatik'), t.get('yer')] if x) or None),
           ('Номинальное напряжение Ue / Ui', ' / '.join(x for x in [t.get('Ue'), t.get('Ui')] if x) or None),
           ('Частота', t.get('f')), ('Система заземления / фазы', ' / '.join(x for x in [t.get('sistem'), t.get('fazlar')] if x) or None),
           ('Номинальный ток НКУ InA', fmt(t.get('In_pano'), ' А')),
           ('Вводной аппарат', f'{gi["poz"]}: {gi["cihaz"]}' if gi else None),
           ('Ток КЗ в точке установки Ik / Icw', ' / '.join(x for x in [fmt(t.get('Ik_proje_kA'), ' кА'), fmt(t.get('Icw_kA'), ' кА (1 с)')] if x) or None),
           ('Напряжение цепей управления', t.get('kumanda')), ('Форма секционирования', t.get('form')),
           ('Ввод / вывод кабелей', k.get('kablo_giris')), ('Материал шин', t.get('bara')),
           ('Нормативный документ', t.get('standart'))]
    rows = [hdr(['Параметр', 'Значение / проектное решение'])]
    zor = {'Габариты В × Ш × Г', 'Степень защиты IP / IK', 'Номинальное напряжение Ue / Ui', 'Система заземления / фазы', 'Номинальный ток НКУ InA'}
    for a, v in par:
        rows.append([para(a, P7B), para(v) if v else Paragraph(('СТОП: ' if a in zor else '') + KTL + ' — не задано', P7R if a in zor else P7A)])
    xT = M + 250; wT = PW - M - xT
    h = doc.table(rows, [170, wT - 170], xT, y)
    yy = y - h - 8
    stops = [s for s in sorun if s[0] == 'STOP']; kt = [s for s in sorun if s[0] == 'KONTROL']
    if stops or kt:
        lines = [f'СТОП: {s[2]}' for s in stops[:4]] + [f'КОНТРОЛЬ: {s[2]}' for s in kt[:max(0, 5 - len(stops[:4]))]]
        more = len(stops) + len(kt) - len(lines)
        if more > 0:
            lines.append(f'… ещё {more} — см. реестр вопросов')
        doc.note_box(xT, yy, wT, 'ДО ЗАПУСКА В ПРОИЗВОДСТВО', lines, RED if stops else AMBER, PALE_RED if stops else PALE_AMB)
    c.setFont('CB', 8); c.setFillColor(INK)
    c.drawString(xT, BOT + 4, 'СОГЛАСОВАНО ЗАКАЗЧИКОМ:  Ф.И.О. ______________________   Подпись ____________   Дата ____________')
    doc.end()


def pano_gabarit(doc, p, sorun):
    c = doc.c; k = p.get('korpus', {})
    doc.page('ПРОИЗВОДСТВЕННОЕ СОГЛАСОВАНИЕ', sub_line(doc.d, p), 'ГАБАРИТНЫЕ И УСТАНОВОЧНЫЕ ВИДЫ', p, durum(sorun))
    H, W, D = k.get('H'), k.get('W'), k.get('D')
    if not (H and W and D):
        c.setFont('CB', 11); c.setFillColor(RED); c.drawString(M, TOP - 30, 'СТОП: габариты корпуса не заданы — виды не построены.'); doc.end(); return
    area_h = TOP - BOT - 60; area_w = 520
    sc = min((area_h) / (H + D + 40), area_w / (W + D + W + 120))
    y_base = BOT + 30
    # ön
    fx = M + 20; sc_, x0, y0, ww, hh = front_sketch(c, fx, y_base, W * sc, H * sc, k, p.get('kapi_elemanlari', []))
    dim_h(c, x0, x0 + ww, y0 - 10, f'{W}'); dim_v(c, x0 - 10, y0, y0 + hh, f'{H}')
    c.setFont('CB', 7); c.drawString(x0, y0 + hh + 8, 'ВИД СПЕРЕДИ — ДВЕРЬ ЗАКРЫТА')
    # yan
    sx = x0 + ww + 50; c.setFillColor(PANEL); c.setStrokeColor(INK); c.rect(sx, y0, D * sc, hh, fill=1)
    dim_h(c, sx, sx + D * sc, y0 - 10, f'{D}'); c.setFont('CB', 7); c.setFillColor(INK); c.drawString(sx, y0 + hh + 8, 'ВИД СБОКУ')
    # üst / alt
    tx = sx + D * sc + 50
    for i, (lab, glands) in enumerate([('ВИД СВЕРХУ', 'сверх' in (k.get('kablo_giris') or '').lower()),
                                       ('ВИД СНИЗУ — КАБЕЛЬНЫЙ ВВОД', 'сниз' in (k.get('kablo_giris') or '').lower() or not k.get('kablo_giris'))]):
        yy = y0 + hh - D * sc - i * (D * sc + 40)
        c.setFillColor(PANEL); c.setStrokeColor(INK); c.rect(tx, yy, W * sc, D * sc, fill=1)
        if glands:
            for j in range(3):
                c.setFillColor(colors.white); c.circle(tx + W * sc * (j + 1) / 4, yy + D * sc / 2, min(6, D * sc / 4), fill=1)
        dim_h(c, tx, tx + W * sc, yy - 10, f'{W}'); c.setFont('CB', 7); c.setFillColor(INK); c.drawString(tx, yy + D * sc + 5, lab)
    # bilgi tablosu
    xT = M + 560; wT = PW - M - xT
    rows = [hdr(['Параметр', 'Значение'])]
    for a, v in [('Установка', k.get('montaj')), ('Дверь', k.get('kapi')), ('Замок', k.get('kilit')),
                 ('Кабельный ввод', k.get('kablo_giris')), ('PE двери', k.get('pe_kapi') or 'Гибкий проводник ПВ3 6 мм² (жёлто-зелёный)'),
                 ('Отверстия ввода', 'Выполнять после подтверждения диаметров кабелей'), ('Цвет', f'RAL {k["RAL"]}' if k.get('RAL') else None)]:
        rows.append([para(a, P7B), para(v) if v else Paragraph(KTL, P7A)])
    h = doc.table(rows, [80, wT - 80], xT, TOP)
    kapi = p.get('kapi_elemanlari', [])
    if kapi:
        rows = [hdr(['Поз.', 'Элемент на двери', 'Цвет', 'Надпись'])]
        for e in kapi:
            rc = renk_kontrol(e)
            rows.append([para(e['poz'], P7B), para(e.get('cihaz')), para(e.get('renk')) if e.get('renk') else Paragraph(KTL, P7A),
                         Paragraph(esc(e.get('yazi') or ''), P7R if rc and rc[0] == 'STOP' else P7)])
        doc.table(rows, [30, wT - 150, 50, 70], xT, TOP - h - 10)
    c.setFont('C', 6.5); c.setFillColor(GREY)
    c.drawString(M, BOT + 4, 'Масштаб условный, пропорции соблюдены. Размеры в мм. Не является чертежом для изготовления корпуса.')
    doc.end()


def pano_yerlesim(doc, p, sorun):
    c = doc.c; k = p.get('korpus', {}); devs = p.get('devreler', [])
    doc.page('ПРОИЗВОДСТВЕННОЕ СОГЛАСОВАНИЕ', sub_line(doc.d, p), 'ВНУТРЕННЯЯ КОМПОНОВКА (ПРЕДВАРИТЕЛЬНАЯ)', p, durum(sorun))
    H, W = k.get('H') or 800, k.get('W') or 600
    ah, aw = TOP - BOT - 30, 470
    sc = min(ah / H, aw / W); ww, hh = W * sc, H * sc; x0 = M + (aw - ww) / 2; y0 = BOT + 18 + (ah - hh) / 2
    c.setStrokeColor(INK); c.setFillColor(PANEL); c.rect(x0, y0, ww, hh, fill=1)
    c.setFillColor(colors.white); c.rect(x0 + 8, y0 + 8, ww - 16, hh - 16, fill=1)
    c.setFont('CB', 7); c.setFillColor(INK); c.drawString(x0, y0 + hh + 6, 'ВИД СПЕРЕДИ — ДВЕРЬ ОТКРЫТА, МОНТАЖНАЯ ПАНЕЛЬ')
    inner_x, inner_w = x0 + 14, ww - 28
    ytop = y0 + hh - 14
    # üst zon: din dışı bloklar
    blok = [d for d in devs if modul(d) is None]
    if blok:
        bw = min(70, inner_w / max(1, min(len(blok), 6)) - 4); bh = 26
        for i, d in enumerate(blok):
            r_, col_ = divmod(i, max(1, int(inner_w // (bw + 4))))
            bx = inner_x + col_ * (bw + 4); by = ytop - bh - r_ * (bh + 4)
            if by < y0 + 40:
                break
            c.setFillColor(colors.HexColor('#DDE3EA')); c.setStrokeColor(INK); c.rect(bx, by, bw, bh, fill=1)
            c.setFillColor(INK)
            yaz(c, bx + bw / 2, by + bh - 9, d['poz'], 'CB', 6, bw - 3)
            yaz(c, bx + bw / 2, by + 10, re.sub(r'\s+', ' ', d.get('cihaz') or ''), 'C', 5, bw - 3)
            yaz(c, bx + bw / 2, by + 3, f'{fmt(d.get("In")) or ""}A' if d.get('In') else '', 'C', 5, bw - 3)
        rows_used = (len(blok) - 1) // max(1, int(inner_w // (bw + 4))) + 1
        ytop -= rows_used * (bh + 4) + 8
    # din sıraları
    din = [d for d in devs if modul(d) is not None]
    nrow = int(k.get('din_sira') or 0); per = float(k.get('sira_modul') or 0)
    mw = (inner_w / per) if per else 9
    if not nrow and din:
        per = per or max(12, int(inner_w // 9)); mw = inner_w / per
        tot = sum(modul(d) for d in din); nrow = int(-(-tot // per))
    gap = 36
    ycur = ytop - 6
    i = 0
    for r in range(nrow):
        if ycur - 28 < y0 + 40:
            c.setFont('CB', 6.5); c.setFillColor(RED); c.drawString(inner_x, ycur - 8, 'Не помещается — см. реестр (R05)'); break
        c.setStrokeColor(STEEL); c.setLineWidth(2); c.line(inner_x, ycur - 14, inner_x + inner_w, ycur - 14); c.setLineWidth(0.5)
        used = 0
        while i < len(din) and used + modul(din[i]) <= per + 1e-6:
            d = din[i]; m_ = modul(d); bx = inner_x + used * mw
            fc = {'mcb': '#FFFFFF', 'rccb': '#E3F2FD', 'rcbo': '#E8F5E9', 'opn': '#FFF3E0'}.get(d['kategori'], '#FFFFFF')
            c.setFillColor(colors.HexColor(fc)); c.setStrokeColor(INK); c.rect(bx, ycur - 26, m_ * mw, 24, fill=1)
            c.setFillColor(INK)
            lab = f'{d.get("egri") or ""}{fmt(d.get("In")) or ""}'
            kw = m_ * mw - 1.5                       # kutu iç genişliği
            if kw < pdfmetrics.stringWidth(d['poz'], 'CB', 3.6):
                # dar (1 modül) kutu: yazı dikey, kutu boyunca
                c.saveState(); c.translate(bx + m_ * mw / 2 + 1.5, ycur - 14); c.rotate(90)
                yaz(c, 0, 0, f'{d["poz"]} {lab}'.strip(), 'CB', 4.4, 22, 'c', 3.2)
                c.restoreState()
            else:
                yaz(c, bx + m_ * mw / 2, ycur - 9, d['poz'], 'CB', 4.8, kw)
                yaz(c, bx + m_ * mw / 2, ycur - 17, lab, 'C', 4.4, kw)
                if d.get('rcd_mA'):
                    yaz(c, bx + m_ * mw / 2, ycur - 23, f'{d["rcd_mA"]}mA', 'C', 4.4, kw)
            used += m_; i += 1
        c.setFont('C', 5.5); c.setFillColor(GREY); c.drawRightString(inner_x + inner_w, ycur - 34, f'ряд {r + 1}: {fmt(used)}/{fmt(per)} мод.')
        ycur -= gap + 8
    # N / PE
    c.setFillColor(colors.HexColor('#1565C0')); c.rect(inner_x, y0 + 16, inner_w * 0.42, 6, fill=1)
    c.setFillColor(colors.HexColor('#F2B705')); c.rect(inner_x + inner_w * 0.55, y0 + 16, inner_w * 0.42, 6, fill=1)
    c.setFillColor(colors.HexColor('#2E7D32'))
    sx_ = inner_x + inner_w * 0.55
    while sx_ < inner_x + inner_w * 0.97 - 6:
        c.rect(sx_, y0 + 16, 4, 6, fill=1, stroke=0); sx_ += 8
    c.setFillColor(INK); c.setFont('CB', 6); c.drawString(inner_x, y0 + 25, 'N'); c.drawString(inner_x + inner_w * 0.55, y0 + 25, 'PE')
    # sağ panel
    xT = M + 490; wT = PW - M - xT; y = TOP
    cap = (int(k.get('din_sira') or 0) * float(k.get('sira_modul') or 0)) or None
    used = sum(modul(d) or 0 for d in devs) + float(k.get('ek_modul') or 0)
    rows = [hdr(['Заполнение', 'Значение'])]
    rows += [[para('Модульные аппараты', P7B), para(f'{fmt(used)} мод.')],
             [para('Ёмкость корпуса', P7B), val_or_ktl(f'{k.get("din_sira")} ряд × {fmt(k.get("sira_modul"))} мод. = {fmt(cap)}' if cap else None)],
             [para('Резерв (норма TEKLED ≥ 20%)', P7B), Paragraph(f'{100 * (cap - used) / cap:.0f}%', P7R if (cap - used) / cap < 0.2 else P7) if cap else Paragraph(KTL, P7A)],
             [para('Немодульные аппараты', P7B), para(', '.join(d['poz'] for d in blok) or '—')]]
    h = doc.table(rows, [130, wT - 130], xT, y); y -= h + 8
    req = ['1. Силовые цепи и цепи управления прокладывать раздельно, в кабель-каналах.',
           '2. Дверь соединить с PE гибким проводником 6 мм².',
           '3. Свободный доступ к аппаратам и клеммам; резерв места по таблице.',
           '4. Маркировка аппаратов и проводов — по таблице линий (лист «Таблица линий»).',
           '5. N и PE не объединять после точки разделения PEN.',
           '6. Порядок аппаратов на DIN-рейке — слева направо по номерам линий.']
    for n_ in p.get('notlar', []):
        req.append('• ' + n_)
    hb = doc.note_box(xT, y, wT, 'МОНТАЖНЫЕ ТРЕБОВАНИЯ', req, GREY, ZEBRA); y -= hb + 8
    if any(s[0] == 'STOP' for s in sorun):
        doc.note_box(xT, y, wT, 'СТОП-ПУНКТ', ['НЕ ЗАПУСКАТЬ РЕЗКУ / СБОРКУ ДО ЗАКРЫТИЯ ВОПРОСОВ ПО РЕЕСТРУ.'])
    c.setFont('C', 6.5); c.setFillColor(GREY)
    c.drawString(M, BOT + 4, 'Схема размещения для согласования. Не является шаблоном сверления; привязки — после проверки паспортов аппаратов.')
    doc.end()


def pano_devre(doc, p, sorun):
    devs = p.get('devreler', [])
    sorunlu = {s[3] for s in sorun if s[3]}
    hd = hdr(['Поз.', 'Аппарат', 'Артикул', 'Марка', 'In, А', 'Полюс', 'Хар.', 'Icu, кА', 'УЗО IΔn / тип', 'Линия / нагрузка', 'P, кВт', 'Кабель', 'Маркировка (текст)'])
    rows = []
    sir = sorted(devs, key=lambda d: (d.get('grup') != 'giris', d.get('kategori') not in ('rubilnik', 'mccb'), d.get('sira', 999)))
    for d in sir:
        cik = d.get('grup') == 'cikis'
        yuk = d.get('yuk') or ('ВВОД' if d.get('grup') == 'giris' else '')
        rcd = f'{d["rcd_mA"]} мА / {d.get("rcd_tip") or "?"}' if d.get('rcd_mA') else ('—' if d.get('kategori') not in ('rccb', 'rcbo') else None)
        stl = P7R if d['poz'] in sorunlu else P7B
        rows.append([para(d['poz'], stl), para(d.get('cihaz')), para(d.get('artikul')), para(d.get('marka')),
                     val_or_ktl(d.get('In')) if d.get('kategori') in ('mcb', 'mccb', 'rcbo', 'rccb', 'rubilnik', 'kontaktor') else para(fmt(d.get('In')) or '—'),
                     para(d.get('kutup') or '—'), para(d.get('egri') or '—'), para(fmt(d.get('Icu_kA')) or '—'),
                     para(rcd) if rcd else Paragraph(KTL, P7A),
                     para(yuk) if yuk or not cik else Paragraph(KTL, P7A), para(fmt(d.get('P_kW')) or '—'),
                     para(d.get('kablo') or '—'), para(f'{d["poz"]} {yuk}'.strip())])
    widths = [32, 150, 52, 44, 30, 32, 26, 34, 52, 120, 30, 82, PW - 2 * M - 684]
    chunks = chunk_rows(hd, rows, widths, TOP - BOT - 24)
    for j, ch in enumerate(chunks):
        doc.page('ПРОИЗВОДСТВЕННОЕ СОГЛАСОВАНИЕ', sub_line(doc.d, p), 'ТАБЛИЦА ЛИНИЙ И МАРКИРОВКИ' + (f' ({j + 1}/{len(chunks)})' if len(chunks) > 1 else ''), p, durum(sorun))
        if not devs:
            doc.c.setFont('CB', 10); doc.c.setFillColor(RED); doc.c.drawString(M, TOP - 20, 'СТОП: аппараты не заданы.')
        else:
            h = doc.table([hd] + ch, widths, M, TOP)
            if j == len(chunks) - 1:
                cnt = Counter(KAT_RU[d['kategori']] for d in devs)
                doc.c.setFont('C', 6.8); doc.c.setFillColor(GREY)
                doc.c.drawString(M, TOP - h - 12, 'Итого: ' + '; '.join(f'{a} — {b} шт' for a, b in cnt.items()) +
                                 '.  Красная позиция — есть вопрос в реестре. Названия линий — по однолинейной схеме проекта.')
        doc.end()


def pano_komple(doc, p, sorun):
    sarf_goster = doc.d.get('genel', {}).get('sarf_goster', False)
    mats = [m for m in p.get('malzemeler', []) if m.get('kategori') not in ('montaj', 'iscilik') and (sarf_goster or m.get('kategori') != 'sarf')]
    hd = hdr(['№', 'Наименование', 'Артикул', 'Марка', 'Кол.', 'Ед.'])
    rows = [[para(i), para(m.get('ad')), para(m.get('artikul')), para(m.get('marka')), para(fmt(m.get('adet'))), para(m.get('birim'))]
            for i, m in enumerate(mats, 1)]
    wT = 520; widths = [22, 270, 80, 70, 38, 40]
    chunks = chunk_rows(hd, rows, widths, TOP - BOT - 30)
    for j, ch in enumerate(chunks):
        doc.page('ПРОИЗВОДСТВЕННОЕ СОГЛАСОВАНИЕ', sub_line(doc.d, p), 'КОМПЛЕКТАЦИЯ (БЕЗ ЦЕН) И КОНТРОЛЬ ПЕРЕД ВЫПУСКОМ' + (f' ({j + 1}/{len(chunks)})' if len(chunks) > 1 else ''), p, durum(sorun))
        c = doc.c
        c.setFont('C', 6.8); c.setFillColor(GREY)
        c.drawString(M, TOP - 2, 'Количество — на 1 комплект, по Excel BOM.' + ('' if sarf_goster else ' Монтажные расходные материалы (шины, провод, клеммы, каналы) — по производственному BOM, в таблицу не включены.'))
        doc.table([hd] + ch, widths, M, TOP - 8)
        if j == 0:
            xR = M + wT + 20; wR = PW - M - xR
            chk = ['Корпус, габариты и IP соответствуют титульному листу', 'Аппараты и артикулы — по таблице линий',
                   'Двери и съёмные металлические части заземлены', 'N и PE выполнены раздельно', 'Момент затяжки силовых соединений проверен',
                   'Сопротивление изоляции и непрерывность PE измерены', 'Маркировка аппаратов, проводов, клемм читаема',
                   'Табличка НКУ и знак «Осторожно! Электрическое напряжение» установлены', 'Механика двери, замка, приводов проверена',
                   'Комплект документации в кармане двери']
            rows2 = [hdr(['☐', 'КОНТРОЛЬ ПЕРЕД ВЫПУСКОМ (ОТК)'])] + [[para('☐'), para(x)] for x in chk]
            doc.table(rows2, [18, wR - 18], xR, TOP - 8)
        doc.end()


def build(data, logo, total=None):
    proje_sorun, sonuc = kontrol_motoru(data)
    # sayfa haritası için ilk geçiş toplamı biliniyorsa da yeniden hesaplanır
    doc = Doc(data, logo, total)
    sayfa_map = data.get('_sayfa_map', {})
    kapak(doc, proje_sorun, sonuc, sayfa_map)
    register(doc, proje_sorun, sonuc)
    genel_gereksinim(doc)
    smap = {}
    for p in data['panolar']:
        s = doc.n + 1
        pano_titul(doc, p, sonuc[p['kod']]); pano_gabarit(doc, p, sonuc[p['kod']]); pano_yerlesim(doc, p, sonuc[p['kod']])
        pano_devre(doc, p, sonuc[p['kod']]); pano_komple(doc, p, sonuc[p['kod']])
        smap[p['kod']] = f'{s}–{doc.n}'
    doc.c.save()
    return doc, smap, proje_sorun, sonuc


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('json_in'); ap.add_argument('pdf_out')
    ap.add_argument('--logo', default='')
    ap.add_argument('--rapor')
    a = ap.parse_args()
    data = json.load(open(a.json_in, encoding='utf-8'))
    doc, smap, _, _ = build(data, a.logo)                       # 1. geçiş: sayfa sayısı + harita
    data['_sayfa_map'] = smap
    doc, smap2, ps, son = build(data, a.logo, total=doc.n)     # 2. geçiş
    assert smap == smap2, 'sayfa haritası değişti'
    open(a.pdf_out, 'wb').write(doc.buf.getvalue())
    tum = ps + [x for v in son.values() for x in v]
    print(f'PDF: {a.pdf_out} | {doc.n} sayfa | STOP {sum(1 for x in tum if x[0] == "STOP")} | KONTROL {sum(1 for x in tum if x[0] == "KONTROL")}')
    print('DURUM:', durum(tum))
    for lev, kod, msg, poz in ps:
        print(f'  [{lev}] Proje {kod}: {msg}')
    for p in data['panolar']:
        for lev, kod, msg, poz in son[p['kod']]:
            print(f'  [{lev}] {p["kod"]} {p["ad"]} {kod}: {msg}')
    if a.rapor:
        json.dump({'durum': durum(tum), 'sayfa': doc.n, 'proje': ps, 'panolar': son}, open(a.rapor, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
