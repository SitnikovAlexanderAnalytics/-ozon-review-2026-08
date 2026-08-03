\
# -*- coding: utf-8 -*-
"""
Строит отчёт ozon-claims-tracker/reports/claims_report.xlsx
Пилотный запуск (Day 0): примеры отобраны через сортировку "Новинки" на Ozon
как удобный (но не полностью надёжный) источник кандидатов.
Начиная со следующего запуска, реальные новинки определяются сравнением
с базой data/products_history.csv (по ID товара).
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

FONT_NAME = "Arial"
TODAY = "2026-07-31"
PROJECT_DIR = r"C:\Users\User\Documents\Работа\Claude Projects\ozon-claims-tracker"
PHOTO_FILES = {
    # Day 0 (2026-07-31) — кандидаты через сортировку "Новинки", не подтверждены сравнением ID
    "5214515666": ("2026-07-31", "5214515666_triol_ponchik.jpg"),
    "5277005073": ("2026-07-31", "5277005073_grandorf_maxi.jpg"),
    "5276044539": ("2026-07-31", "5276044539_award_duck.jpg"),
    "5259070728": ("2026-07-31", "5259070728_grandorf_wet.jpg"),
    "5280647420": ("2026-07-31", "5280647420_clanvet_hepatic.jpg"),
    "4685211777": ("2026-07-31", "4685211777_lopay_lapoy.jpg"),
    "5110846419": ("2026-07-31", "5110846419_alma_kasha.jpg"),
    "5277005334": ("2026-07-31", "5277005334_premier_cat.jpg"),
    "5243322000": ("2026-07-31", "5243322000_edwin.jpg"),
    "5248949561": ("2026-07-31", "5248949561_grandorf_cat.jpg"),
    "4098300679": ("2026-07-31", "4098300679_titbit_trava.jpg"),
    # Day 1 (2026-08-01) — подтверждены сравнением ID с базой (ID выше максимума, известного на 2026-07-31)
    "5293611532": ("2026-08-01", "5293611532_gosbi_gf_puppy.jpg"),
    "5294915094": ("2026-08-01", "5294915094_gosbi_gf_cat.jpg"),
    "5294272736": ("2026-08-01", "5294272736_nyamtoriya_assorti.jpg"),
    "5293523232": ("2026-08-01", "5293523232_murchik.jpg"),
    # Day 2-3 (2026-08-03) — подтверждены сравнением ID с базой (ID выше максимума на 2026-08-01)
    "5306421483": ("2026-08-03", "5306421483_livera_lamb.jpg"),
    "5306617135": ("2026-08-03", "5306617135_livera_hypo.jpg"),
    "5327569955": ("2026-08-03", "5327569955_farmina_hypo.jpg"),
    "5327539554": ("2026-08-03", "5327539554_farmina_renal.jpg"),
    "5327528788": ("2026-08-03", "5327528788_farmina_diabetic.jpg"),
    "5327751575": ("2026-08-03", "5327751575_vashstandart.jpg"),
    "5321602313": ("2026-08-03", "5321602313_delicachoice.jpg"),
    "5294994812": ("2026-08-03", "5294994812_zillii.jpg"),
    "5306618148": ("2026-08-03", "5306618148_livera_cat.jpg"),
    "5259705799": ("2026-08-03", "5259705799_award_cat.jpg"),
    "5254153891": ("2026-08-03", "5254153891_grandorf_cat2.jpg"),
    "5322942235": ("2026-08-03", "5322942235_kasha_generic.jpg"),
}

def photo_path(oid):
    date, fname = PHOTO_FILES[oid]
    return os.path.join(PROJECT_DIR, "photos", date, fname)

TARGET_H_PX = 130

wb = openpyxl.Workbook()

# ---------- Стили ----------
header_fill = PatternFill(start_color="2F5233", end_color="2F5233", fill_type="solid")
header_font = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
title_font = Font(name=FONT_NAME, bold=True, size=14)
note_font = Font(name=FONT_NAME, italic=True, size=9, color="666666")
base_font = Font(name=FONT_NAME, size=10)
wrap = Alignment(wrap_text=True, vertical="top")
thin = Side(style="thin", color="CCCCCC")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = border

def autosize(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

# =====================================================================
# Sheet 1: Примеры
# =====================================================================
ws = wb.active
ws.title = "Примеры"

ws["A1"] = "Клеймы новинок — Корма и лакомства для собак и кошек (Ozon)"
ws["A1"].font = title_font
ws["A2"] = ("2026-07-31 (Day 0): кандидаты отобраны через фильтр Ozon «Новинки» — быстрый, но не полностью "
            "надёжный источник (Ozon продвигает не только объективно новые товары). "
            "2026-08-01 и 2026-08-03 (Day 1, Day 2-3): товары подтверждены объективно — сравнением ID с базой "
            "data/products_history.csv; их ID выше максимума, зафиксированного в предыдущем снимке той же категории. "
            "Колонки «Протеин» и «Ингредиенты-клеймы» — только то, что вынесено крупно на лицевую сторону упаковки "
            "(не полный состав мелким шрифтом). «Завод-изготовитель» — юрлицо и адрес производства с этикетки "
            "(не бренд и не импортёр/представитель), где это удалось разобрать по фото.")
ws["A2"].font = note_font
ws.merge_cells("A2:N2")
ws["A2"].alignment = wrap
ws.row_dimensions[2].height = 58

headers = ["Дата обнаружения", "Метод обнаружения", "Вид животного", "Тип продукции", "Бренд", "Название товара",
           "Протеин", "Ингредиенты-клеймы (фасад)", "Завод-изготовитель", "ID на Ozon", "Ссылка", "Фото упаковки",
           "Кол-во категорий клеймов", "Все клеймы (кратко)"]
hrow = 4
for i, h in enumerate(headers, start=1):
    ws.cell(row=hrow, column=i, value=h)
style_header_row(ws, hrow, len(headers))

products = [
    dict(animal="Собаки", ptype="Лакомства", brand="Triol",
         name="Пончик жевательный с мясом утки",
         oid="5214515666",
         url="https://www.ozon.ru/product/zhevatelnyy-ponchik-s-myasom-utki-dlya-sobak-5214515666/",
         photo="https://ir.ozone.ru/s3/multimedia-1-n/wc1000/13008229403.jpg",
         claims_short="100% натурально; Dental (зубной налёт, дёсны); привлекает внимание",
         protein="Утка", ingredients="—",
         plant="Не найдено (в галерее Ozon только предметные фото товара, без этикетки состава)"),
    dict(animal="Собаки", ptype="Сухой корм", brand="GRANDORF",
         name="Fresh Adult MAXI, ягнёнок с бататом, 3 кг (беззерновой)",
         oid="5277005073",
         url="https://www.ozon.ru/product/suhoy-bezzernovoy-korm-grandorf-fresh-adult-maxi-dlya-sobak-krupnyh-porod-yagnenok-s-batatom-3-kg-v-5277005073/",
         photo="https://ir.ozone.ru/s3/multimedia-1-q/wc1000/13183424738.jpg",
         claims_short="Holistic&Hypoallergenic; Gut Health/пробиотики; глюкозамин/хондроитин; Omega-3/6; Grain Free; мясо ягнёнка первым ингредиентом",
         protein="Ягнёнок", ingredients="Батат; Omega-3/Omega-6; глюкозамин и хондроитин; живые пробиотики",
         plant="United Petfood Producer NV, Бельгия (завод PET 010); также указаны заводы группы во Франции и Нидерландах"),
    dict(animal="Собаки", ptype="Сухой корм", brand="AWARD",
         name="Monoprotein, утка с яблоком и артишоком, 12 кг",
         oid="5276044539",
         url="https://www.ozon.ru/product/suhoy-korm-dlya-sobak-vseh-porod-award-monoprotein-iz-utki-s-dobavleniem-yablok-i-artishoka-12kg-5276044539/",
         photo="https://ir.ozone.ru/s3/multimedia-1-f/wc1000/13221055959.jpg",
         claims_short="Premium Quality; Monoprotein; Hypo&Digestion; Low Grain; 100% BIO без ГМО; поддержка ЖКТ/иммунитета/кожи/мочевыдел.системы",
         protein="Утка", ingredients="Яблоко; артишок",
         plant="ООО «ПетЭксперт», Липецкая обл., Елецкий р-н, ОЭЗ ППТ «Липецк»; произведено для АО «Валта Пет Продактс» (Москва)"),
    dict(animal="Собаки", ptype="Влажный корм", brand="GRANDORF",
         name="Fresh, паштет из индейки, беззерновой, 6х200 г",
         oid="5259070728",
         url="https://www.ozon.ru/product/vlazhnyy-korm-dlya-sobak-s-indeykoy-bezzernovoy-pashtet-grandorf-konservy-gipoallergennye-nabor-5259070728/",
         photo="https://ir.ozone.ru/s3/multimedia-1-e/wc1000/13158897062.jpg",
         claims_short="Holistic&Hypoallergenic; 70% мяса; Monoprotein diet; Grain Free",
         protein="Индейка", ingredients="—",
         plant="Не проверено отдельно; вероятно United Petfood (тот же концерн, что и GRANDORF Fresh Adult MAXI)"),
    dict(animal="Собаки", ptype="Влажный корм", brand="CLAN VET",
         name="Hepatic — профилактика болезней печени, 340 г",
         oid="5280647420",
         url="https://www.ozon.ru/product/vlazhnyy-dieticheskiy-korm-clan-vet-hepatic-dlya-sobak-profilaktika-bolezney-pecheni-340g-5280647420/",
         photo="https://ir.ozone.ru/s3/multimedia-1-k/wc1000/13235339612.jpg",
         claims_short="VET / ветеринарная профилактическая диета; профилактика болезней печени",
         protein="Курица (субпродукты, по составу)", ingredients="—",
         plant="г. Елец, Липецкая обл., Россия (низкая уверенность в статусе — вероятно тот же промышленный кластер, что и ООО «ПетЭксперт» у AWARD)"),
    dict(animal="Собаки", ptype="Лакомства", brand="Лопай лапой",
         name="Лёгкое говяжье, сублимированное (freeze dried)",
         oid="4685211777",
         url="https://www.ozon.ru/product/lakomstvo-sublimirovannoe-legkoe-govyazhe-naturalnye-freeze-dried-dog-treats-premialnyy-perekus-i-4685211777/",
         photo="https://ir.ozone.ru/s3/multimedia-1-b/wc1000/11144536823.jpg",
         claims_short="100% натуральный вкус; технология Freeze Dried; способствует очищению зубов; удобно для поощрения",
         protein="Говядина", ingredients="—",
         plant="Не найдено (все 8 фото в галерее — маркетинговые инфографики без этикетки состава)"),
    dict(animal="Собаки", ptype="Каши", brand="Альма",
         name="Каша с говядиной, быстрого приготовления",
         oid="5110846419",
         url="https://www.ozon.ru/product/kasha-s-govyadinoy-alma-korm-dlya-sobak-naturalnyy-5110846419/",
         photo="https://ir.ozone.ru/s3/multimedia-v/wc1000/6858505111.jpg",
         claims_short="Не требует варки; натуральное мясо в составе",
         protein="Говядина", ingredients="—",
         plant="ИП [фамилия неразборчива], г. Волгоград, Россия (адрес читается, имя предпринимателя — нет)"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Premier",
         name="Lamb & Turkey Adult Large, гипоаллергенный, 400 г",
         oid="5277005334",
         url="https://www.ozon.ru/product/gipoallergennyy-suhoy-korm-dlya-vzroslyh-koshek-krupnyh-porod-premier-cat-lamb-turkey-adult-5277005334/",
         photo="https://ir.ozone.ru/s3/multimedia-1-i/wc1000/13183412922.jpg",
         claims_short="Hypoallergenic; при чувствит. пищеварении; Low Grain/Gluten Free; без курицы/кукурузы/пшеницы/сои; мясо первым ингредиентом; Made in Europe",
         protein="Ягнёнок + индейка", ingredients="—",
         plant="Не проверено отдельно; на фасаде заявлено Made in Europe (страна не детализирована)"),
    dict(animal="Кошки", ptype="Сухой корм", brand="EDWIN",
         name="Для кошек, склонных к аллергии, индейка, 8 кг",
         oid="5243322000",
         url="https://www.ozon.ru/product/korm-suhoy-edwin-dlya-koshek-sklonnyh-k-allergii-indeyka-8-kg-5243322000/",
         photo="https://ir.ozone.ru/s3/multimedia-1-n/wc1000/13105121819.jpg",
         claims_short="Premium Food for Cats; Hypoallergenic; оптимальное пищеварение; защита иммунитета; высокое содержание белка",
         protein="Индейка", ingredients="—",
         plant="ООО «ПроРацион», Россия, 630110, г. Новосибирск, ул. Писемского, д. 24/1"),
    dict(animal="Кошки", ptype="Влажный корм", brand="GRANDORF",
         name="Fresh, филе тунца с лососем в желе, 6х70 г",
         oid="5248949561",
         url="https://www.ozon.ru/product/grandorf-fresh-vlazhnyy-korm-dlya-koshek-tunets-losos-6sht-h-70g-5248949561/",
         photo="https://ir.ozone.ru/s3/multimedia-1-q/wc1000/7453596482.jpg",
         claims_short="Holistic Nutrition; 60% мяса; филе тунца/лосося; Сделано в Таиланде",
         protein="Тунец + лосось", ingredients="—",
         plant="Таиланд (страна заявлена на фасаде, конкретное юрлицо не проверялось)"),
    dict(animal="Кошки", ptype="Лакомства", brand="Titbit",
         name="Трава для кошек (для проращивания), 40 г х 12",
         oid="4098300679",
         url="https://www.ozon.ru/product/titbit-lakomstva-dlya-koshek-oves-dlya-prorashchivaniya-40-gr-12-up-4098300679/",
         photo="https://ir.ozone.ru/s3/multimedia-5/wc1000/6878482385.jpg",
         claims_short="100% натурально; немодифицированные семена; способствует пищеварению/выведению шерсти",
         protein="—(не мясной продукт, трава)", ingredients="Овёс; витамины и фолиевая кислота (восполняет запас)",
         plant="Не найдено (только 2 фото в галерее, без этикетки)"),
    # --- Day 1 (2026-08-01): подтверждено сравнением ID с базой (реальные новинки, не сортировка) ---
    dict(animal="Собаки", ptype="Сухой корм", brand="Gosbi",
         name="PROFESSIONAL GF Puppy, лосось с ягнёнком, 18 кг (беззерновой)",
         oid="5293611532",
         url="https://www.ozon.ru/product/gosbi-professional-gf-puppy-losos-yagnenok-18-kg-suhoy-bezzernovoy-korm-dlya-shchenkov-vseh-porod-5293611532/",
         photo="https://ir.ozone.ru/s3/multimedia-1-r/wc1000/13281279255.jpg",
         claims_short="Grain Free; Mediterranean gourmet; Professional (премиум-позиционирование)",
         protein="Лосось + ягнёнок", ingredients="—",
         plant="Не проверено отдельно; Gosbi — испанский бренд (Средиземноморское позиционирование на упаковке)"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Gosbi",
         name="PROFESSIONAL GF Cat Adult, курица с индейкой, 18 кг (беззерновой)",
         oid="5294915094",
         url="https://www.ozon.ru/product/gosbi-professional-gf-cat-adult-kuritsa-indeyka-18-kg-suhoy-bezzernovoy-korm-dlya-koshek-5294915094/",
         photo="https://ir.ozone.ru/s3/multimedia-1-1/wc1000/13284455941.jpg",
         claims_short="Grain Free; Mediterranean gourmet; Professional (премиум-позиционирование)",
         protein="Курица + индейка", ingredients="—",
         plant="Не проверено отдельно; Gosbi — испанский бренд (Средиземноморское позиционирование на упаковке)"),
    dict(animal="Собаки", ptype="Лакомства", brand="НЯМТОРИЯ",
         name="Ассорти «Мясное разнообразие 2.0» (утка/тунец/курица), 3х80 г",
         oid="5294272736",
         url="https://www.ozon.ru/product/assorti-lakomstv-dlya-sobak-nyamtoriya-myasnoe-raznoobrazie-2-0-240-g-3-vida-80-g-h-3-5294272736/",
         photo="https://ir.ozone.ru/s3/multimedia-1-9/wc1000/13281800301.jpg",
         claims_short="Без глицерина (утка/курица; вкус тунец глицерин СОДЕРЖИТ — не универсально для линейки); 90% мяса (74,1% в тунце); стейки с водорослями",
         protein="Утка / тунец+курица+треска / курица (3 вкуса в ассорти)", ingredients="Морская капуста (водоросли); морковь",
         plant="Не найдено (проверено 7 из 14 фото в галерее — все по-вкусовые инфографики, без этикетки с адресом)"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Мурчик",
         name="Говядина в соусе, кусочки, 75 г х 35 шт (эконом-сегмент)",
         oid="5293523232",
         url="https://www.ozon.ru/product/murchik-vlazhnyy-korm-dlya-koshek-govyadina-v-souse-75g-35sht-5293523232/",
         photo="https://ir.ozone.ru/s3/multimedia-1-i/wc1000/13278946986.jpg",
         claims_short="Нет заявленных клеймов — только базовое описание (кусочки в соусе)",
         protein="Говядина", ingredients="—",
         plant="Не проверено (эконом-сегмент, вероятно РФ)"),
    # --- Day 2-3 (2026-08-03): подтверждено сравнением ID, репрезентативная выборка из 34 найденных ---
    dict(animal="Собаки", ptype="Сухой корм", brand="LiveRa",
         name="Adult Lamb & Rice, ягнёнок с рисом, 3,5 кг",
         oid="5306421483",
         url="https://www.ozon.ru/product/livera-suhoy-korm-dlya-sobak-adult-lamb-rice-yagnenok-s-risom-3-5kg-5306421483/",
         photo="https://ir.ozone.ru/s3/multimedia-1-z/wc1000/13322588867.jpg",
         claims_short="100% nature; Holistic",
         protein="Ягнёнок", ingredients="Рис",
         plant="г. Барнаул, Россия (бренд заявляет «Корма из России»; юрлицо на этикетке нечитаемо)"),
    dict(animal="Собаки", ptype="Сухой корм", brand="LiveRa",
         name="Hypoallergenic, индейка с рисом, для щенков и собак мелких пород, 7 кг",
         oid="5306617135",
         url="https://www.ozon.ru/product/livera-suhoy-korm-dlya-shchenkov-i-sobak-melkih-porod-hypoallergenic-indeyka-s-risom-7kg-5306617135/",
         photo="https://ir.ozone.ru/s3/multimedia-1-1/wc1000/13323355381.jpg",
         claims_short="Hypoallergenic; при пищевой аллергии и чувствит. пищеварении; 61% мяса; 100% natural; Holistic",
         protein="Индейка", ingredients="Рис",
         plant="г. Барнаул, Россия (тот же бренд LiveRa)"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Farmina VetLife",
         name="Hypoallergenic Duck & Potato, диетический, 6х300 г",
         oid="5327569955",
         url="https://www.ozon.ru/product/dieticheskiy-vlazhnyy-korm-dlya-sobak-farmina-vet-life-natural-diet-dog-hypoallergenic-duck-5327569955/",
         photo="https://ir.ozone.ru/s3/multimedia-1-b/wc1000/13389282695.jpg",
         claims_short="При пищевой аллергии; и непереносимости продуктов; с низким риском развития аллергий; Farmina Pet Foods с 1965 года",
         protein="Утка", ingredients="Картофель",
         plant="Farmina Pet Foods, Италия (established 1965 — вероятно сам производитель, не аутсорс; отдельно не проверялось)"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Farmina VetLife",
         name="Renal — диета при заболеваниях почек, 6х300 г",
         oid="5327539554",
         url="https://www.ozon.ru/product/dieticheskiy-vlazhnyy-korm-dlya-sobak-farmina-vet-life-natural-diet-dog-renal-6x300gr-5327539554/",
         photo="https://ir.ozone.ru/s3/multimedia-1-u/wc1000/13389041874.jpg",
         claims_short="Поддерживает здоровье почек; снижает нагрузку на почки; подходит для пожилых собак",
         protein="Не указан на фасаде", ingredients="—",
         plant="Farmina Pet Foods, Италия"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Farmina VetLife",
         name="Diabetic — диета для собак с диабетом, 6х300 г",
         oid="5327528788",
         url="https://www.ozon.ru/product/dieticheskiy-vlazhnyy-korm-dlya-sobak-farmina-vet-life-natural-diet-dog-diabetic-300h6-5327528788/",
         photo="https://ir.ozone.ru/s3/multimedia-1-l/wc1000/13388901681.jpg",
         claims_short="Поддерживает здоровое пищеварение; контроль уровня глюкозы в крови; подходит для собак с диабетом",
         protein="Не указан на фасаде", ingredients="—",
         plant="Farmina Pet Foods, Италия"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Ваш Стандарт",
         name="Для мелких пород, утка, премиум, 600 г",
         oid="5327751575",
         url="https://www.ozon.ru/product/suhoy-korm-dlya-sobak-vash-standart-dlya-melkih-porod-utka-premium-polnoratsionnyy-600-gr-5327751575/",
         photo="https://ir.ozone.ru/s3/multimedia-1-k/wc1000/13210223780.jpg",
         claims_short="ГОСТ; «без ГМО и переплат»; Сделано на Кубани",
         protein="Утка (по названию; на стоковом фото изображена говядина — вероятно, шаблонная картинка бренда)",
         ingredients="—",
         plant="ИП Клюкова А.А., Краснодарский край, Приморско-Ахтарский р-н («Сделано на Кубани»)"),
    dict(animal="Собаки", ptype="Лакомства", brand="DelicaChoice",
         name="Кости из жил прессованные, с говядиной, 5 см",
         oid="5321602313",
         url="https://www.ozon.ru/product/lakomstvo-delicachoice-dlya-sobak-kost-5-sm-5321602313/",
         photo="https://ir.ozone.ru/s3/multimedia-1-0/wc1000/13377006936.jpg",
         claims_short="Нет заявленных клеймов — базовый жевательный продукт (говяжья кожа)",
         protein="Говядина (говяжья кожа)", ingredients="—",
         plant="Тяньцзинь Гэ Шен Пэт Продактс Ко., ЛТД, Тяньцзинь, Китай; импортёр ООО «Евро-Продукт-К», г. Краснодар"),
    dict(animal="Кошки", ptype="Сухой корм", brand="ZILLII",
         name="Light/Sterilized, белая рыба с лососем, 10 кг",
         oid="5294994812",
         url="https://www.ozon.ru/product/suhoy-korm-dlya-sterilizovannyh-koshek-s-izbytochnym-vesom-zillii-zillii-light-sterilized-cat-belaya-5294994812/",
         photo="https://ir.ozone.ru/s3/multimedia-1-7/wc1000/13284776239.jpg",
         claims_short="SUPER PREMIUM CAT FOOD; NO SOY/WHEAT/CORN/GMO/colourants added; эко-инициатива «посади дерево»",
         protein="Белая рыба + лосось", ingredients="—",
         plant="Не найдено (не проверялась галерея дополнительно)"),
    dict(animal="Кошки", ptype="Сухой корм", brand="LiveRa",
         name="Sensitive, чувствительное пищеварение, индейка с рисом, 0,7 кг",
         oid="5306618148",
         url="https://www.ozon.ru/product/livera-suhoy-korm-dlya-koshek-sensitive-s-chuvstvitelnym-pishchevareniem-indeyka-s-risom-0-7kg-5306618148/",
         photo="https://ir.ozone.ru/s3/multimedia-1-p/wc1000/13323368581.jpg",
         claims_short="Чувствительное пищеварение; 100% natural; Hypoallergenic; Holistic; 67% мяса; «Корма из России»",
         protein="Индейка", ingredients="Рис",
         plant="Россия (прямо заявлено «Корма из России» на упаковке; тот же бренд LiveRa, вероятно г. Барнаул)"),
    dict(animal="Кошки", ptype="Влажный корм", brand="AWARD",
         name="Паштет из индейки с овощами, консервы, 100 г х 12 шт",
         oid="5259705799",
         url="https://www.ozon.ru/product/award-vlazhnyy-korm-dlya-vzroslyh-koshek-pashtet-iz-indeyki-s-ovoshchami-v-konservah-100-g-h-12-sht-5259705799/",
         photo="https://ir.ozone.ru/s3/multimedia-1-t/wc1000/7093491617.jpg",
         claims_short="Monoprotein — Single Protein Source; 100% BIO",
         protein="Индейка", ingredients="Овощи",
         plant="Не проверено отдельно; вероятно тот же завод, что у AWARD dog (ООО «ПетЭксперт», Липецкая обл.)"),
    dict(animal="Кошки", ptype="Влажный корм", brand="GRANDORF",
         name="Куриная грудка с лососем, беззерновой, консервы, 6х70 г",
         oid="5254153891",
         url="https://www.ozon.ru/product/vlazhnyy-korm-grandorf-konservy-bezzernovoy-dlya-lyubyh-koshek-kurinaya-grudka-s-lososem-6sht-h-70g-5254153891/",
         photo="https://ir.ozone.ru/s3/multimedia-1-7/wc1000/7423644211.jpg",
         claims_short="NATURAL & HEALTHY; GRAIN FREE; Holistic Nutrition for Cats",
         protein="Курица + лосось", ingredients="—",
         plant="Не проверено отдельно для этого SKU; см. другие GRANDORF (United Petfood/Таиланд)"),
    dict(animal="Собаки", ptype="Каши", brand="СибКорм",
         name="Каша для собак с мясом, 10 кг",
         oid="5322942235",
         url="https://www.ozon.ru/product/korm-dlya-sobak-kasha-s-myasom-10-kg-5322942235/",
         photo="https://ir.ozone.ru/s3/multimedia-1-i/wc1000/12755719170.jpg",
         claims_short="Произведено на Алтае; 100% натуральный состав; без ГМО; из натуральных ингредиентов; для всех пород и возрастов",
         protein="Мясо (вид не конкретизирован на фасаде)", ingredients="—",
         plant="Алтайский край, Россия («Произведено на Алтае» — прямо на упаковке)"),
]

# категория-код на товар (для подсчёта количества категорий на товар)
cats_per_product = [3, 5, 6, 3, 1, 4, 2, 5, 4, 3, 2,  2, 2, 3, 0,
                     1, 3, 3, 2, 2, 2, 0, 1, 3, 2, 2, 2]

METHOD_BY_DATE = {
    "2026-07-31": "Сортировка «Новинки» (не подтверждено ID)",
    "2026-08-01": "Подтверждено: ID выше макс. известного на 2026-07-31",
    "2026-08-03": "Подтверждено: ID выше макс. известного на 2026-08-01",
}
PHOTO_COL = 12

r = hrow + 1
max_photo_w = 0
for p, ncat in zip(products, cats_per_product):
    pdate, pfname = PHOTO_FILES[p["oid"]]
    method = METHOD_BY_DATE[pdate]
    ws.cell(row=r, column=1, value=pdate)
    ws.cell(row=r, column=2, value=method)
    ws.cell(row=r, column=3, value=p["animal"])
    ws.cell(row=r, column=4, value=p["ptype"])
    ws.cell(row=r, column=5, value=p["brand"])
    ws.cell(row=r, column=6, value=p["name"])
    ws.cell(row=r, column=7, value=p["protein"])
    ws.cell(row=r, column=8, value=p["ingredients"])
    ws.cell(row=r, column=9, value=p["plant"])
    ws.cell(row=r, column=10, value=p["oid"])
    cell = ws.cell(row=r, column=11, value=p["url"])
    cell.hyperlink = p["url"]
    cell.font = Font(name=FONT_NAME, size=10, color="0563C1", underline="single")
    ws.cell(row=r, column=13, value=ncat)
    ws.cell(row=r, column=14, value=p["claims_short"])
    for c in range(1, 15):
        if c != 11:
            ws.cell(row=r, column=c).font = base_font
        ws.cell(row=r, column=c).alignment = wrap
        ws.cell(row=r, column=c).border = border

    # Встроенное превью фото упаковки (не ссылка, а сама картинка)
    ppath = photo_path(p["oid"])
    im = PILImage.open(ppath)
    w0, h0 = im.size
    target_w = int(w0 * TARGET_H_PX / h0)
    max_photo_w = max(max_photo_w, target_w)
    img = XLImage(ppath)
    img.height = TARGET_H_PX
    img.width = target_w
    ws.add_image(img, f"{get_column_letter(PHOTO_COL)}{r}")
    ws.row_dimensions[r].height = TARGET_H_PX * 0.78

    r += 1

last_row = r - 1
autosize(ws, [13, 30, 11, 12, 14, 34, 22, 34, 34, 12, 10,
              round(max_photo_w / 7) + 1, 10, 46])
ws.freeze_panes = "A5"

# =====================================================================
# Sheet 2: Клеймы (детализация: 1 строка = товар x категория)
# =====================================================================
ws2 = wb.create_sheet("Клеймы")
headers2 = ["ID товара", "Название товара", "Вид животного", "Тип продукции",
            "Код категории", "Категория клейма", "Формулировки на упаковке"]
for i, h in enumerate(headers2, start=1):
    ws2.cell(row=1, column=i, value=h)
style_header_row(ws2, 1, len(headers2))

claim_rows = [
    ("5214515666", "Пончик жевательный с мясом утки", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натурально"),
    ("5214515666", "Пончик жевательный с мясом утки", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Dental; снимает зубной налёт; массирует дёсна"),
    ("5214515666", "Пончик жевательный с мясом утки", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведение", "Привлекают внимание собаки на долгое время"),

    ("5277005073", "GRANDORF Fresh Adult MAXI", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Holistic & Hypoallergenic; при чувствительном пищеварении"),
    ("5277005073", "GRANDORF Fresh Adult MAXI", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Gut Health — Live Probiotics; Glucosamine and Chondroitin; Omega-3, Omega-6 Fatty Acids"),
    ("5277005073", "GRANDORF Fresh Adult MAXI", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% Natural Dog Food"),
    ("5277005073", "GRANDORF Fresh Adult MAXI", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free; No: Grains, Chicken, By-products"),
    ("5277005073", "GRANDORF Fresh Adult MAXI", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Fresh Lamb is the first ingredient"),

    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypo & Digestion — при пищевой аллергии и непереносимости"),
    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Поддержка пищеварения и иммунитета; поддержка мочевыделяющей системы; здоровая кожа и блестящая шерсть; крепкий иммунитет каждый день"),
    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% BIO; без ГМО, красителей и ароматизаторов"),
    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Monoprotein; Low Grain"),
    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Fresh Meat"),
    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Premium Quality; Superpremium Dog Food"),

    ("5259070728", "GRANDORF Fresh (индейка, консервы)", "Собаки", "Влажный корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Holistic & Hypoallergenic"),
    ("5259070728", "GRANDORF Fresh (индейка, консервы)", "Собаки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Monoprotein diet; Grain Free"),
    ("5259070728", "GRANDORF Fresh (индейка, консервы)", "Собаки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "70% Meat Ingredients"),

    ("5280647420", "CLAN VET Hepatic", "Собаки", "Влажный корм", "VET", "Ветеринарная / лечебная диета", "VET; Hepatic — профилактика болезней печени; профилактическая диета"),

    ("4685211777", "Лопай лапой — говяжье сублимированное", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральный вкус"),
    ("4685211777", "Лопай лапой — говяжье сублимированное", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Способствует очищению зубов"),
    ("4685211777", "Лопай лапой — говяжье сублимированное", "Собаки", "Лакомства", "TECH", "Особая технология производства", "Инновационная технология Freeze Dried; сублимированное"),
    ("4685211777", "Лопай лапой — говяжье сублимированное", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведение", "Удобно для поощрения; аппетитные кусочки; любимое лакомство с первого кусочка"),

    ("5110846419", "Альма — каша с говядиной", "Собаки", "Каши", "NAT", "Натуральность состава", "Каша с мясом (говядина)"),
    ("5110846419", "Альма — каша с говядиной", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки; быстрого приготовления"),

    ("5277005334", "Premier Lamb & Turkey Adult Large", "Кошки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic; при чувствительном пищеварении; No chicken, corn, wheat, soy"),
    ("5277005334", "Premier Lamb & Turkey Adult Large", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Procare — профилактика здоровья каждый день"),
    ("5277005334", "Premier Lamb & Turkey Adult Large", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Low Grain; Gluten Free"),
    ("5277005334", "Premier Lamb & Turkey Adult Large", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "1 ingredient is deboned Fresh Meat"),
    ("5277005334", "Premier Lamb & Turkey Adult Large", "Кошки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Made in Europe"),

    ("5243322000", "EDWIN — для кошек, склонных к аллергии", "Кошки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic"),
    ("5243322000", "EDWIN — для кошек, склонных к аллергии", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Оптимальное пищеварение; защита иммунитета"),
    ("5243322000", "EDWIN — для кошек, склонных к аллергии", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Высокое содержание белка"),
    ("5243322000", "EDWIN — для кошек, склонных к аллергии", "Кошки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Premium Food for Cats"),

    ("5248949561", "GRANDORF Fresh — тунец с лососем", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Holistic Nutrition for Cats"),
    ("5248949561", "GRANDORF Fresh — тунец с лососем", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "60% Meat; филе тунца и лосося"),
    ("5248949561", "GRANDORF Fresh — тунец с лососем", "Кошки", "Влажный корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Сделано в Таиланде"),

    ("4098300679", "Titbit — трава для кошек", "Кошки", "Лакомства", "NAT", "Натуральность состава", "100% натурально; немодифицированные семена"),
    ("4098300679", "Titbit — трава для кошек", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Способствует пищеварению / выведению шерсти"),

    # --- Day 1 (2026-08-01), подтверждено сравнением ID ---
    ("5293611532", "Gosbi PROFESSIONAL GF Puppy", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5293611532", "Gosbi PROFESSIONAL GF Puppy", "Собаки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Mediterranean gourmet; Professional"),

    ("5294915094", "Gosbi PROFESSIONAL GF Cat Adult", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free / Getreidefrei"),
    ("5294915094", "Gosbi PROFESSIONAL GF Cat Adult", "Кошки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Mediterranean gourmet; Professional"),

    ("5294272736", "НЯМТОРИЯ — Мясное разнообразие 2.0", "Собаки", "Лакомства", "NAT", "Натуральность состава", "Без глицерина"),
    ("5294272736", "НЯМТОРИЯ — Мясное разнообразие 2.0", "Собаки", "Лакомства", "MEAT", "Происхождение и качество мяса", "90% мяса; 74,1% мяса рыбы (тунец)"),
    ("5294272736", "НЯМТОРИЯ — Мясное разнообразие 2.0", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Стейки с добавлением водорослей"),

    # --- Day 2-3 (2026-08-03), подтверждено сравнением ID ---
    ("5306421483", "LiveRa Adult Lamb & Rice", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% nature; Holistic"),

    ("5306617135", "LiveRa Hypoallergenic Turkey & Rice", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic; при пищевой аллергии и чувствительном пищеварении"),
    ("5306617135", "LiveRa Hypoallergenic Turkey & Rice", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% natural; Holistic"),
    ("5306617135", "LiveRa Hypoallergenic Turkey & Rice", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "61% мяса"),

    ("5327569955", "Farmina VetLife Hypoallergenic Duck&Potato", "Собаки", "Влажный корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "При пищевой аллергии; с низким риском развития аллергий"),
    ("5327569955", "Farmina VetLife Hypoallergenic Duck&Potato", "Собаки", "Влажный корм", "VET", "Ветеринарная / лечебная диета", "Vet Life диетический корм; и непереносимости продуктов"),
    ("5327569955", "Farmina VetLife Hypoallergenic Duck&Potato", "Собаки", "Влажный корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Farmina Pet Foods — с 1965 года"),

    ("5327539554", "Farmina VetLife Renal", "Собаки", "Влажный корм", "VET", "Ветеринарная / лечебная диета", "Диета при заболеваниях почек; подходит для пожилых собак"),
    ("5327539554", "Farmina VetLife Renal", "Собаки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Поддерживает здоровье почек; снижает нагрузку на почки"),

    ("5327528788", "Farmina VetLife Diabetic", "Собаки", "Влажный корм", "VET", "Ветеринарная / лечебная диета", "Диета для собак с диабетом"),
    ("5327528788", "Farmina VetLife Diabetic", "Собаки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Поддерживает здоровое пищеварение; контроль уровня глюкозы в крови"),

    ("5327751575", "Ваш Стандарт (утка)", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "Без ГМО и переплат"),
    ("5327751575", "Ваш Стандарт (утка)", "Собаки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "ГОСТ; Сделано на Кубани"),

    ("5294994812", "ZILLII Light/Sterilized (белая рыба+лосось)", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "NO SOY/WHEAT/CORN/GMO/colourants added"),

    ("5306618148", "LiveRa Sensitive Turkey&Rice (кошки)", "Кошки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Чувствительное пищеварение; Hypoallergenic"),
    ("5306618148", "LiveRa Sensitive Turkey&Rice (кошки)", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% natural; Holistic; «Корма из России»"),
    ("5306618148", "LiveRa Sensitive Turkey&Rice (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "67% мяса"),

    ("5259705799", "AWARD Turkey Pate (кошки)", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Monoprotein — Single Protein Source"),
    ("5259705799", "AWARD Turkey Pate (кошки)", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "100% BIO"),

    ("5254153891", "GRANDORF Cat курица+лосось", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5254153891", "GRANDORF Cat курица+лосось", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Natural & Healthy; Holistic Nutrition for Cats"),

    ("5322942235", "СибКорм — каша с мясом", "Собаки", "Каши", "NAT", "Натуральность состава", "100% натуральный состав; без ГМО; из натуральных ингредиентов"),
    ("5322942235", "СибКорм — каша с мясом", "Собаки", "Каши", "ORIGIN", "Страна производства / премиум-позиционирование", "Произведено на Алтае"),
]

for i, row in enumerate(claim_rows, start=2):
    for j, val in enumerate(row, start=1):
        c = ws2.cell(row=i, column=j, value=val)
        c.font = base_font
        c.alignment = wrap
        c.border = border

autosize(ws2, [12, 30, 12, 12, 10, 32, 55])
ws2.freeze_panes = "A2"

# =====================================================================
# Sheet 3: Клеймы-справочник (таксономия)
# =====================================================================
ws3 = wb.create_sheet("Клеймы-справочник")
headers3 = ["Код", "Категория клейма", "Описание", "Типичные формулировки"]
for i, h in enumerate(headers3, start=1):
    ws3.cell(row=1, column=i, value=h)
style_header_row(ws3, 1, len(headers3))

taxonomy = [
    ("HYPO", "Гипоаллергенность / чувствит. пищеварение",
     "Заявления об отсутствии типичных аллергенов, подходит при непереносимости, для чувствительного ЖКТ",
     "Hypoallergenic; при чувствительном пищеварении; Hypo & Digestion; без курицы/кукурузы/пшеницы/сои"),
    ("FUNC", "Функциональная польза для здоровья",
     "Конкретные функциональные преимущества сверх базового питания: иммунитет, суставы, кожа/шерсть, мочевыдел. система, ЖКТ",
     "Поддержка иммунитета; Glucosamine & Chondroitin; Omega-3/6; Gut Health; здоровая кожа и шерсть"),
    ("DENT", "Стоматологическая польза",
     "Уход за зубами и дёснами",
     "Dental; снимает зубной налёт; массирует дёсна; способствует очищению зубов"),
    ("VET", "Ветеринарная / лечебная диета",
     "Диета для профилактики или коррекции конкретных заболеваний, ветеринарное позиционирование бренда",
     "VET; профилактика болезней печени/почек/ЖКТ; лечебный корм"),
    ("NAT", "Натуральность состава",
     "Отсутствие искусственных добавок, ГМО, «чистая» этикетка",
     "100% натурально; без ГМО, красителей и ароматизаторов; 100% BIO; Holistic Nutrition"),
    ("GRAIN", "Беззерновой / монобелковый состав",
     "Особый тип рецептуры по составу углеводов/белка",
     "Grain Free; Monoprotein; Low Grain; Gluten Free"),
    ("MEAT", "Происхождение и качество мяса",
     "Доля мяса, место мяса в составе, вид/качество ингредиента",
     "Мясо первым ингредиентом; 60-70% мяса; филе тунца/лосося; высокое содержание белка"),
    ("TECH", "Особая технология производства",
     "Способ обработки продукта, отличный от стандартного",
     "Freeze Dried; сублимированное; холодное прессование"),
    ("CONV", "Удобство использования",
     "Практичность для владельца при приготовлении/хранении",
     "Не требует варки; готовый рацион; быстрого приготовления"),
    ("ORIGIN", "Страна производства / премиум-позиционирование",
     "Указание страны изготовления или premium/super-premium статуса",
     "Made in Europe; Сделано в Таиланде; Premium Quality; Superpremium"),
    ("TASTE", "Вкусовая привлекательность / поведенческая польза",
     "Ориентация на удовольствие питомца или поведенческий эффект (игра, дрессировка, вывод шерсти)",
     "Любимое лакомство; для дрессировки; привлекает внимание; трава для кошек"),
]
for i, row in enumerate(taxonomy, start=2):
    for j, val in enumerate(row, start=1):
        c = ws3.cell(row=i, column=j, value=val)
        c.font = base_font
        c.alignment = wrap
        c.border = border

autosize(ws3, [9, 32, 55, 55])
ws3.freeze_panes = "A2"

# =====================================================================
# Sheet 4: Аналитика
# =====================================================================
ws4 = wb.create_sheet("Аналитика")
ws4["A1"] = "Аналитика: популярность направлений клеймов среди новинок"
ws4["A1"].font = title_font
ws4["A2"] = (f"Выборка: {len(products)} карточек (собаки+кошки, корма и лакомства) за 3 снимка — "
             "11 отобраны 2026-07-31 через сортировку «Новинки» (не подтверждено ID), "
             "4 добавлены 2026-08-01 и 12 добавлены 2026-08-03 — все 16 подтверждены сравнением ID с базой "
             "(объективно новые; ещё 22 объективно новых SKU за 2026-08-03 не вошли в выборку — см. заметку в выводах). "
             "Каждая категория считается один раз на товар (не на фразу).")
ws4["A2"].font = note_font
ws4.merge_cells("A2:D2")
ws4["A2"].alignment = wrap
ws4.row_dimensions[2].height = 30

headers4 = ["Код", "Категория клейма", "Товаров с этой категорией", "Доля от выборки"]
hrow4 = 4
for i, h in enumerate(headers4, start=1):
    ws4.cell(row=hrow4, column=i, value=h)
style_header_row(ws4, hrow4, len(headers4))

codes = [t[0] for t in taxonomy]
names = {t[0]: t[1] for t in taxonomy}
n_products = len(products)

r = hrow4 + 1
first_data_row = r
for code in codes:
    ws4.cell(row=r, column=1, value=code)
    ws4.cell(row=r, column=2, value=names[code])
    # COUNTIF по коду категории в листе "Клеймы" — каждая строка там = 1 товар x 1 категория,
    # поэтому обычный COUNTIF корректно даёт число уникальных товаров с этой категорией.
    ws4.cell(row=r, column=3, value=f'=COUNTIF(Клеймы!E:E,A{r})')
    ws4.cell(row=r, column=4, value=f'=C{r}/{n_products}')
    ws4.cell(row=r, column=4).number_format = "0%"
    for c in range(1, 5):
        ws4.cell(row=r, column=c).font = base_font
        ws4.cell(row=r, column=c).border = border
    r += 1
last_data_row = r - 1

ws4.cell(row=r, column=1, value="Итого товаров в выборке:")
ws4.cell(row=r, column=1).font = Font(name=FONT_NAME, bold=True, size=10)
ws4.merge_cells(f"A{r}:B{r}")
ws4.cell(row=r, column=3, value=n_products)
ws4.cell(row=r, column=3).font = Font(name=FONT_NAME, bold=True, size=10)

autosize(ws4, [9, 34, 24, 14])

# Диаграмма — популярность категорий клеймов
from openpyxl.chart import BarChart, Reference
chart = BarChart()
chart.type = "bar"
chart.title = "Товаров-новинок с данной категорией клейма (из {})".format(n_products)
chart.y_axis.title = None
chart.x_axis.title = "Количество товаров"
data = Reference(ws4, min_col=3, min_row=hrow4, max_row=last_data_row)
cats = Reference(ws4, min_col=2, min_row=first_data_row, max_row=last_data_row)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.height = 10
chart.width = 20
ws4.add_chart(chart, f"F{hrow4}")

# ---- Выводы ----
concl_row = last_data_row + 22
ws4.cell(row=concl_row, column=1, value="Выводы (наблюдения: 2026-07-31, 2026-08-01, 2026-08-03)").font = Font(name=FONT_NAME, bold=True, size=12)
conclusions = [
    "1. Натуральность состава (NAT) — по-прежнему самое частое направление и только укрепилось: 16 из 27 карточек "
    "(59%) заявляют «100% натурально», «без ГМО», «без глицерина» или Holistic-позиционирование. Это «входной "
    "билет» для нового SKU почти в любом сегменте и ценовой категории, а не способ выделиться.",
    "2. Происхождение мяса (MEAT) и страна-производитель/премиум (ORIGIN) выросли до 9 из 27 (33%) каждая — но "
    "ORIGIN теперь работает на ДВЕ разные аудитории одновременно: премиальное «Made in Europe / Mediterranean "
    "gourmet / Сделано в Таиланде» (Gosbi, GRANDORF, Premier) и патриотично-региональное «Сделано на Кубани», "
    "«Произведено на Алтае», «Корма из России» (Ваш Стандарт, СибКорм, LiveRa) — один и тот же код клейма "
    "скрывает два противоположных позиционирования: глобальный премиум и локальный патриотизм.",
    "3. НОВОЕ: ветеринарная/лечебная диета (VET) выросла с 1 до 4 из 27 (15%) за счёт единовременного запуска "
    "линейки Farmina VetLife (Hypoallergenic, Renal, Diabetic) — это уже не единичный случай, а признак того, что "
    "лечебное питание становится самостоятельным растущим направлением, а не нишей одного бренда.",
    "4. Гипоаллергенность (HYPO) и функциональная польза (FUNC) держатся на 8 из 27 (30%) каждая — обе всё так же "
    "характерны для сухих кормов средне-премиального и премиального сегмента (GRANDORF, AWARD, Premier, EDWIN, "
    "LiveRa, Farmina), и практически не встречаются в лакомствах и эконом-сегменте.",
    "5. Завод-изготовитель вскрыл географический кластер контрактного производства в России: ООО «ПетЭксперт» "
    "(Липецкая обл., Елецкий р-н, ОЭЗ «Липецк») делает корм под брендом AWARD для АО «Валта Пет Продактс» и, "
    "вероятно, тот же кластер — под CLAN VET; ООО «ПроРацион» (Новосибирск) делает EDWIN; бренд LiveRa указывает "
    "«Корма из России» с адресом в Барнауле. Как и с GRANDORF/United Petfood в Европе, один и тот же завод может "
    "стоять за несколькими разными брендами на полке Ozon — бренд и реальный изготовитель систематически различны.",
    "6. Протеиновый разрыв премиум/эконом: индейка и утка — основные протеины у гипоаллергенных и "
    "премиальных SKU (EDWIN, LiveRa Hypoallergenic, Farmina Duck, Premier, GRANDORF) — как менее аллергенные и "
    "более «премиальные» виды мяса; говядина доминирует в эконом-сегменте и базовых лакомствах (Мурчик, "
    "DelicaChoice, Лопай лапой, Альма) — как самое доступное и привычное мясное сырьё.",
    "7. Сегментный разрыв в клеймах подтверждён на двух независимых эконом-товарах подряд: «Мурчик» (влажный "
    "корм для кошек) и «DelicaChoice» (жевательная кость из говяжьей кожи, Китай) — оба объективно новые SKU, "
    "и ни один не несёт ни одного клейма из справочника, только вес/вкус/базовое описание. Messaging-борьба "
    "почти полностью сосредоточена в премиум- и средне-премиальном сегменте.",
    "8. Важный нюанс по составу: клейм не всегда универсален внутри одной товарной линейки. У НЯМТОРИЯ «без "
    "глицерина» подтверждается для вкусов утка и курица, но вкус тунец, по составу на этикетке, глицерин "
    "СОДЕРЖИТ — «чистая этикетка» может относиться только к части ассортимента бренда, а не ко всей линейке.",
    "",
    "Методология: 11 карточек за 2026-07-31 отобраны через сортировку Ozon «Новинки» и НЕ подтверждены "
    "объективно (это подтвердилось на практике: при следующих заходах состав «Новинок» почти полностью "
    "обновлялся, включая товары со значительно более старыми ID, чем уже известные, — фильтр Ozon явно "
    "алгоритмический, не хронологический). 16 карточек (4 за 2026-08-01, 12 за 2026-08-03) прошли объективную "
    "проверку: их ID выше максимума, зафиксированного в той же подкатегории на момент предыдущего снимка, — то "
    "есть они физически не могли существовать раньше. За 2026-08-03 такой проверке объективно соответствовали "
    "34 SKU (из них 12 разобраны подробно с фото и клеймами; ещё 22 — в основном вкусовые/весовые вариации той "
    "же линейки одного из этих 12 брендов — учтены только в data/products_history.csv, без отдельного разбора "
    "клеймов). Устойчивая недельная динамика появится после нескольких дней накопления базы "
    "data/products_history.csv. Колонки «Протеин» и «Ингредиенты-клеймы» отражают только то, что вынесено "
    "крупно на лицевую сторону упаковки — не полный состав мелким шрифтом. «Завод-изготовитель» указан там, "
    "где адрес и юрлицо на фото удалось разобрать; для части товаров в галерее Ozon нет фото с этикеткой "
    "состава вовсе (отмечено как «не найдено»).",
]
rr = concl_row + 1
for line in conclusions:
    ws4.cell(row=rr, column=1, value=line)
    ws4.cell(row=rr, column=1).font = base_font if line else base_font
    ws4.cell(row=rr, column=1).alignment = wrap
    ws4.merge_cells(f"A{rr}:K{rr}")
    ws4.row_dimensions[rr].height = (56 if len(line) > 260 else 42) if line else 8
    rr += 1

wb.save(r"C:\Users\User\Documents\Работа\Claude Projects\ozon-claims-tracker\reports\claims_report_v2.xlsx")
print("saved")
