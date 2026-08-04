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
    # Доразбор 01.08 (ранее полностью пропущенные)
    "5294201871": ("2026-08-01", "5294201871_gosbi_puppymini.jpg"),
    "5294241487": ("2026-08-01", "5294241487_gosbi_puppy.jpg"),
    "5294353414": ("2026-08-01", "5294353414_gosbi_adultmini_lamb.jpg"),
    "5294317994": ("2026-08-01", "5294317994_nyamtoriya_maxi.jpg"),
    "5294208319": ("2026-08-01", "5294208319_nyamtoriya_meatvariety.jpg"),
    "5294159811": ("2026-08-01", "5294159811_nyamtoriya_duck.jpg"),
    "5294111572": ("2026-08-01", "5294111572_nyamtoriya_mini.jpg"),
    "5294843939": ("2026-08-01", "5294843939_gosbi_kitten.jpg"),
    "5294960856": ("2026-08-01", "5294960856_gosbi_sterilised.jpg"),
    # Доразбор 03.08 (ранее только протеин без клеймов/завода)
    "5306470815": ("2026-08-03", "5306470815_livera_fish.jpg"),
    "5306440073": ("2026-08-03", "5306440073_livera_poultry.jpg"),
    "5306424879": ("2026-08-03", "5306424879_livera_rabbit.jpg"),
    "5308822504": ("2026-08-03", "5308822504_houndhall_cartilage.jpg"),
    "5324402513": ("2026-08-03", "5324402513_kingkorm_rumen.jpg"),
    "5321600961": ("2026-08-03", "5321600961_dognis_bone.jpg"),
    "5328154887": ("2026-08-03", "5328154887_allcats.jpg"),
    "5325606109": ("2026-08-03", "5325606109_mealgrand_turkeychicken.jpg"),
    "5325020749": ("2026-08-03", "5325020749_mealgrand_lamb.jpg"),
    "5254153857": ("2026-08-03", "5254153857_grandorf_chicken.jpg"),
    "5254153733": ("2026-08-03", "5254153733_grandorf_tunashrimp.jpg"),
    "5254153795": ("2026-08-03", "5254153795_grandorf_tunacrab.jpg"),
    "5327393201": ("2026-08-03", "5327393201_vashstandart_assort.jpg"),
    "5327553656": ("2026-08-03", "5327553656_farmina_gastro.jpg"),
    "5259703870": ("2026-08-03", "5259703870_award_sterilized.jpg"),
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

    # ---- Доразбор пропущенных SKU (01.08 без разбора + 03.08 без клеймов/завода) ----
    dict(animal="Собаки", ptype="Сухой корм", brand="Gosbi",
         name="PROFESSIONAL Puppy Mini (Курица, Рыба), 18 кг",
         oid="5294201871",
         url="https://www.ozon.ru/product/5294201871/",
         photo="https://ir.ozone.ru/s3/multimedia-1-v/wc1000/13318572271.jpg",
         claims_short="Mediterranean gourmet; Made in Spain",
         protein="Курица, Рыба", ingredients="—", plant="Испания"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Gosbi",
         name="PROFESSIONAL Puppy (Курица, Рыба), 18 кг, для щенков всех пород",
         oid="5294241487",
         url="https://www.ozon.ru/product/5294241487/",
         photo="https://ir.ozone.ru/s3/multimedia-1-5/wc1000/13319182589.jpg",
         claims_short="Mediterranean gourmet; Made in Spain",
         protein="Курица, Рыба", ingredients="—", plant="Испания"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Gosbi",
         name="PROFESSIONAL Adult Mini Lamb & Rice (Ягнёнок, рис), 18 кг",
         oid="5294353414",
         url="https://www.ozon.ru/product/5294353414/",
         photo="https://ir.ozone.ru/s3/multimedia-1-5/wc1000/13321406309.jpg",
         claims_short="Mediterranean gourmet; Made in Spain",
         protein="Ягнёнок", ingredients="Рис", plant="Испания"),
    dict(animal="Собаки", ptype="Лакомства", brand="НЯМТОРИЯ",
         name='Ассорти "Вкусняшки макси 2.0", 210 г (стейки ягнёнок/утка/карпаччо утка)',
         oid="5294317994",
         url="https://www.ozon.ru/product/5294317994/",
         photo="https://ir.ozone.ru/s3/multimedia-1-m/wc1000/13281985858.jpg",
         claims_short="Без глицерина; натуральное мясо премиум-качества; 92%/74%/100% мяса",
         protein="Ягнёнок, Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="НЯМТОРИЯ",
         name='Ассорти "Мясное разнообразие", 240 г (говядина/ягнёнок/курица с водорослями)',
         oid="5294208319",
         url="https://www.ozon.ru/product/5294208319/",
         photo="https://ir.ozone.ru/s3/multimedia-1-z/wc1000/13281563627.jpg",
         claims_short="Без глицерина; натуральное мясо премиум-качества; 92%/90% мяса",
         protein="Говядина, Ягнёнок, Курица", ingredients="Морская капуста (водоросли)", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="НЯМТОРИЯ",
         name='Ассорти "Деликатесы из утки", 210 г (мраморные стейки/карпаччо/стейки с водорослями)',
         oid="5294159811",
         url="https://www.ozon.ru/product/5294159811/",
         photo="https://ir.ozone.ru/s3/multimedia-1-2/wc1000/13281392702.jpg",
         claims_short="Без глицерина; натуральное мясо премиум-качества; 94%/100%/90% мяса",
         protein="Утка", ingredients="Морская капуста (водоросли)", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="НЯМТОРИЯ",
         name='Ассорти "Вкусняшки мини 2.0", 240 г (говядина/утка с водорослями/курица с водорослями)',
         oid="5294111572",
         url="https://www.ozon.ru/product/5294111572/",
         photo="https://ir.ozone.ru/s3/multimedia-1-2/wc1000/13281177386.jpg",
         claims_short="Без глицерина; натуральное мясо премиум-качества; 88%/90%/90% мяса",
         protein="Говядина, Утка, Курица", ingredients="Морская капуста (водоросли)", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Gosbi",
         name="PROFESSIONAL Cat Kitten (Курица, Индейка), 18 кг, для котят и берем./кормящих кошек",
         oid="5294843939",
         url="https://www.ozon.ru/product/5294843939/",
         photo="https://ir.ozone.ru/s3/multimedia-1-n/wc1000/13322357915.jpg",
         claims_short="Mediterranean gourmet; Made in Spain",
         protein="Курица, Индейка", ingredients="—", plant="Испания"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Gosbi",
         name="PROFESSIONAL Cat Sterilised (Курица, Лосось), 18 кг",
         oid="5294960856",
         url="https://www.ozon.ru/product/5294960856/",
         photo="https://ir.ozone.ru/s3/multimedia-1-w/wc1000/13321853960.jpg",
         claims_short="Mediterranean gourmet; Made in Spain",
         protein="Курица, Лосось", ingredients="—", plant="Испания"),
    dict(animal="Собаки", ptype="Сухой корм", brand="LiveRa",
         name="Adult Fish, филе белой рыбы с рисом, 3,5 кг",
         oid="5306470815",
         url="https://www.ozon.ru/product/5306470815/",
         photo="https://ir.ozone.ru/s3/multimedia-1-u/wc1000/13322751150.jpg",
         claims_short="100% natural Hypoallergenic; 51% мяса",
         protein="Белая рыба", ingredients="Рис", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="LiveRa",
         name="Adult, ассорти из мяса домашней птицы, 10 кг",
         oid="5306440073",
         url="https://www.ozon.ru/product/5306440073/",
         photo="https://ir.ozone.ru/s3/multimedia-1-0/wc1000/13322671092.jpg",
         claims_short="100% natural Hypoallergenic; Holistic; 54% мяса; Супер премиум класс",
         protein="Птица (ассорти)", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="LiveRa",
         name="Adult Large, кролик с рисом, 1,2 кг",
         oid="5306424879",
         url="https://www.ozon.ru/product/5306424879/",
         photo="https://ir.ozone.ru/s3/multimedia-1-r/wc1000/13322595267.jpg",
         claims_short="100% natural Hypoallergenic; Holistic; 63% мяса; Корма из России; супер премиум класса",
         protein="Кролик", ingredients="Рис", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="ХАУНДХОЛЛ",
         name="Лопаточный хрящ, большой размер, 1 шт",
         oid="5308822504",
         url="https://www.ozon.ru/product/5308822504/",
         photo="https://ir.ozone.ru/s3/multimedia-1-c/wc1000/13324599264.jpg",
         claims_short="100% натурально",
         protein="Не указан на фасаде", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="KingKorm",
         name="Рубец бараний, 100 г",
         oid="5324402513",
         url="https://www.ozon.ru/product/5324402513/",
         photo="https://ir.ozone.ru/s3/multimedia-1-u/wc1000/13386364914.jpg",
         claims_short="100% натурально; без красителей и ГМО; поддерживает здоровье зубов; улучшает пищеварение; 100% баранина",
         protein="Баранина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="DOGNIS",
         name='Кости из жил прессованные 8" (20 см), 1 шт',
         oid="5321600961",
         url="https://www.ozon.ru/product/5321600961/",
         photo="https://ir.ozone.ru/s3/multimedia-1-w/wc1000/13377004376.jpg",
         claims_short="Без заявленных клеймов",
         protein="Не указан на фасаде", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="All Cats (AllerPetfood)",
         name="Sterilized Cats, с индейкой, 13 кг",
         oid="5328154887",
         url="https://www.ozon.ru/product/5328154887/",
         photo="https://ir.ozone.ru/s3/multimedia-1-6/wc1000/7435761450.jpg",
         claims_short="Complete premium pet food; Датское качество; Omega-3/6; Taurine; Dental Care",
         protein="Индейка", ingredients="Омега-3/6", plant="Дания (страна бренда, не завод)"),
    dict(animal="Кошки", ptype="Сухой корм", brand="MEALGRAND",
         name="Супер Премиум, для стерилизованных кошек, с индейкой и курицей, 8 кг",
         oid="5325606109",
         url="https://www.ozon.ru/product/5325606109/",
         photo="https://ir.ozone.ru/s3/multimedia-1-m/wc1000/13390392010.jpg",
         claims_short="Супер Премиум",
         protein="Индейка, Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="MEALGRAND",
         name="Супер Премиум, для стерилизованных кошек, с ягнёнком, 8 кг",
         oid="5325020749",
         url="https://www.ozon.ru/product/5325020749/",
         photo="https://ir.ozone.ru/s3/multimedia-1-2/wc1000/13388477186.jpg",
         claims_short="Супер Премиум",
         protein="Ягнёнок", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="GRANDORF",
         name="Cat, с куриной грудкой, 70 г х 6 шт",
         oid="5254153857",
         url="https://www.ozon.ru/product/5254153857/",
         photo="https://ir.ozone.ru/s3/multimedia-1-r/wc1000/7519254327.jpg",
         claims_short="Grain Free; Natural & Healthy; Holistic Nutrition for Cats",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="GRANDORF",
         name="Cat, филе тунца с креветками, 70 г х 6 шт",
         oid="5254153733",
         url="https://www.ozon.ru/product/5254153733/",
         photo="https://ir.ozone.ru/s3/multimedia-1-s/wc1000/7119962668.jpg",
         claims_short="Grain Free; Natural & Healthy; Holistic Nutrition for Cats",
         protein="Тунец, Креветки", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="GRANDORF",
         name="Cat, филе тунца с мясом краба, 70 г х 6 шт",
         oid="5254153795",
         url="https://www.ozon.ru/product/5254153795/",
         photo="https://ir.ozone.ru/s3/multimedia-1-h/wc1000/7119962657.jpg",
         claims_short="Grain Free; Natural & Healthy; Holistic Nutrition for Cats",
         protein="Тунец, Краб", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Ваш Стандарт",
         name="Мясное ассорти, премиум, 4 кг, для средних/крупных пород",
         oid="5327393201",
         url="https://www.ozon.ru/product/5327393201/",
         photo="https://ir.ozone.ru/s3/multimedia-1-d/wc1000/13168813837.jpg",
         claims_short="Без пафоса и переплат; Сделано на Кубани; ГОСТ",
         protein="Не конкретизирован (мясное ассорти)", ingredients="—",
         plant="ИП Клюкова А.А., Краснодарский край, Приморско-Ахтарский р-н"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Farmina VetLife",
         name="Gastrointestinal, 6х300г",
         oid="5327553656",
         url="https://www.ozon.ru/product/5327553656/",
         photo="https://ir.ozone.ru/s3/multimedia-1-i/wc1000/13389160770.jpg",
         claims_short="Здоровое ЖКТ; поддержка пищеварения; сильный иммунитет",
         protein="Не указан на фасаде", ingredients="—", plant="Farmina Pet Foods, Италия"),
    dict(animal="Кошки", ptype="Влажный корм", brand="AWARD",
         name="Sterilized, кусочки в соусе с индейкой, 85г х12",
         oid="5259703870",
         url="https://www.ozon.ru/product/5259703870/",
         photo="https://ir.ozone.ru/s3/multimedia-1-8/wc1000/8018668556.jpg",
         claims_short="100% BIO; Urinary Health; Nutraceutical Formula; Superpremium Cat Food",
         protein="Индейка", ingredients="—", plant="Не найдено"),
]

# категория-код на товар (для подсчёта количества категорий на товар)
cats_per_product = [3, 6, 6, 4, 1, 4, 2, 5, 4, 3, 2,  3, 3, 3, 0,
                     2, 4, 3, 2, 2, 3, 0, 1, 5, 2, 3, 2,
                     # доразбор 01.08 + 03.08 (в порядке добавления в products)
                     1, 1, 1, 3, 3, 3, 3, 1, 1,
                     3, 5, 6, 1, 4, 0, 4, 1, 1, 3, 3, 3, 3, 2, 3]

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

    ("5277005073", "GRANDORF Fresh Adult MAXI", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic; при чувствительном пищеварении"),
    ("5277005073", "GRANDORF Fresh Adult MAXI", "Собаки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
    ("5277005073", "GRANDORF Fresh Adult MAXI", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Gut Health — Live Probiotics; Glucosamine and Chondroitin; Omega-3, Omega-6 Fatty Acids"),
    ("5277005073", "GRANDORF Fresh Adult MAXI", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% Natural Dog Food"),
    ("5277005073", "GRANDORF Fresh Adult MAXI", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free; No: Grains, Chicken, By-products"),
    ("5277005073", "GRANDORF Fresh Adult MAXI", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Fresh Lamb is the first ingredient"),

    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypo & Digestion — при пищевой аллергии и непереносимости"),
    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Поддержка пищеварения и иммунитета; поддержка мочевыделяющей системы; здоровая кожа и блестящая шерсть; крепкий иммунитет каждый день"),
    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% BIO; без ГМО, красителей и ароматизаторов"),
    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Monoprotein; Low Grain"),
    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Fresh Meat"),
    ("5276044539", "AWARD Monoprotein (утка)", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Premium Quality; Superpremium Dog Food"),

    ("5259070728", "GRANDORF Fresh (индейка, консервы)", "Собаки", "Влажный корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic"),
    ("5259070728", "GRANDORF Fresh (индейка, консервы)", "Собаки", "Влажный корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
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
    ("5243322000", "EDWIN — для кошек, склонных к аллергии", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Premium Food for Cats"),

    ("5248949561", "GRANDORF Fresh — тунец с лососем", "Кошки", "Влажный корм", "HOLISTIC", "Холистик-позиционирование", "Holistic Nutrition for Cats"),
    ("5248949561", "GRANDORF Fresh — тунец с лососем", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "60% Meat; филе тунца и лосося"),
    ("5248949561", "GRANDORF Fresh — тунец с лососем", "Кошки", "Влажный корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Сделано в Таиланде"),

    ("4098300679", "Titbit — трава для кошек", "Кошки", "Лакомства", "NAT", "Натуральность состава", "100% натурально; немодифицированные семена"),
    ("4098300679", "Titbit — трава для кошек", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Способствует пищеварению / выведению шерсти"),

    # --- Day 1 (2026-08-01), подтверждено сравнением ID ---
    ("5293611532", "Gosbi PROFESSIONAL GF Puppy", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5293611532", "Gosbi PROFESSIONAL GF Puppy", "Собаки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Mediterranean gourmet"),
    ("5293611532", "Gosbi PROFESSIONAL GF Puppy", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Professional"),

    ("5294915094", "Gosbi PROFESSIONAL GF Cat Adult", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free / Getreidefrei"),
    ("5294915094", "Gosbi PROFESSIONAL GF Cat Adult", "Кошки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Mediterranean gourmet"),
    ("5294915094", "Gosbi PROFESSIONAL GF Cat Adult", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Professional"),

    ("5294272736", "НЯМТОРИЯ — Мясное разнообразие 2.0", "Собаки", "Лакомства", "NAT", "Натуральность состава", "Без глицерина"),
    ("5294272736", "НЯМТОРИЯ — Мясное разнообразие 2.0", "Собаки", "Лакомства", "MEAT", "Происхождение и качество мяса", "90% мяса; 74,1% мяса рыбы (тунец)"),
    ("5294272736", "НЯМТОРИЯ — Мясное разнообразие 2.0", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Стейки с добавлением водорослей"),

    # --- Day 2-3 (2026-08-03), подтверждено сравнением ID ---
    ("5306421483", "LiveRa Adult Lamb & Rice", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% nature"),
    ("5306421483", "LiveRa Adult Lamb & Rice", "Собаки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),

    ("5306617135", "LiveRa Hypoallergenic Turkey & Rice", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic; при пищевой аллергии и чувствительном пищеварении"),
    ("5306617135", "LiveRa Hypoallergenic Turkey & Rice", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% natural"),
    ("5306617135", "LiveRa Hypoallergenic Turkey & Rice", "Собаки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
    ("5306617135", "LiveRa Hypoallergenic Turkey & Rice", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "61% мяса"),

    ("5327569955", "Farmina VetLife Hypoallergenic Duck&Potato", "Собаки", "Влажный корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "При пищевой аллергии; с низким риском развития аллергий"),
    ("5327569955", "Farmina VetLife Hypoallergenic Duck&Potato", "Собаки", "Влажный корм", "VET", "Ветеринарная / лечебная диета", "Vet Life диетический корм; и непереносимости продуктов"),
    ("5327569955", "Farmina VetLife Hypoallergenic Duck&Potato", "Собаки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Farmina Pet Foods — с 1965 года"),

    ("5327539554", "Farmina VetLife Renal", "Собаки", "Влажный корм", "VET", "Ветеринарная / лечебная диета", "Диета при заболеваниях почек; подходит для пожилых собак"),
    ("5327539554", "Farmina VetLife Renal", "Собаки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Поддерживает здоровье почек; снижает нагрузку на почки"),

    ("5327528788", "Farmina VetLife Diabetic", "Собаки", "Влажный корм", "VET", "Ветеринарная / лечебная диета", "Диета для собак с диабетом"),
    ("5327528788", "Farmina VetLife Diabetic", "Собаки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Поддерживает здоровое пищеварение; контроль уровня глюкозы в крови"),

    ("5327751575", "Ваш Стандарт (утка)", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "Без ГМО и переплат"),
    ("5327751575", "Ваш Стандарт (утка)", "Собаки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Сделано на Кубани"),
    ("5327751575", "Ваш Стандарт (утка)", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "ГОСТ"),

    ("5294994812", "ZILLII Light/Sterilized (белая рыба+лосось)", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "NO SOY/WHEAT/CORN/GMO/colourants added"),

    ("5306618148", "LiveRa Sensitive Turkey&Rice (кошки)", "Кошки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Чувствительное пищеварение; Hypoallergenic"),
    ("5306618148", "LiveRa Sensitive Turkey&Rice (кошки)", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% natural"),
    ("5306618148", "LiveRa Sensitive Turkey&Rice (кошки)", "Кошки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
    ("5306618148", "LiveRa Sensitive Turkey&Rice (кошки)", "Кошки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "«Корма из России»"),
    ("5306618148", "LiveRa Sensitive Turkey&Rice (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "67% мяса"),

    ("5259705799", "AWARD Turkey Pate (кошки)", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Monoprotein — Single Protein Source"),
    ("5259705799", "AWARD Turkey Pate (кошки)", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "100% BIO"),

    ("5254153891", "GRANDORF Cat курица+лосось", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5254153891", "GRANDORF Cat курица+лосось", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Natural & Healthy"),
    ("5254153891", "GRANDORF Cat курица+лосось", "Кошки", "Влажный корм", "HOLISTIC", "Холистик-позиционирование", "Holistic Nutrition for Cats"),

    ("5322942235", "СибКорм — каша с мясом", "Собаки", "Каши", "NAT", "Натуральность состава", "100% натуральный состав; без ГМО; из натуральных ингредиентов"),
    ("5322942235", "СибКорм — каша с мясом", "Собаки", "Каши", "ORIGIN", "Страна производства / премиум-позиционирование", "Произведено на Алтае"),

    # ---- Доразбор 01.08 ----
    ("5294201871", "Gosbi Puppy Mini", "Собаки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Mediterranean gourmet; Made in Spain"),
    ("5294241487", "Gosbi Puppy", "Собаки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Mediterranean gourmet; Made in Spain"),
    ("5294353414", "Gosbi Adult Mini Lamb&Rice", "Собаки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Mediterranean gourmet; Made in Spain"),

    ("5294317994", "НЯМТОРИЯ Вкусняшки макси 2.0", "Собаки", "Лакомства", "NAT", "Натуральность состава", "Без глицерина; натуральное мясо премиум-качества"),
    ("5294317994", "НЯМТОРИЯ Вкусняшки макси 2.0", "Собаки", "Лакомства", "MEAT", "Происхождение и качество мяса", "92% мяса / 74% мяса / 100% утка"),
    ("5294317994", "НЯМТОРИЯ Вкусняшки макси 2.0", "Собаки", "Лакомства", "PREMIUM", "Премиальное позиционирование", "мясо премиум-качества"),

    ("5294208319", "НЯМТОРИЯ Мясное разнообразие", "Собаки", "Лакомства", "NAT", "Натуральность состава", "Без глицерина; натуральное мясо премиум-качества"),
    ("5294208319", "НЯМТОРИЯ Мясное разнообразие", "Собаки", "Лакомства", "MEAT", "Происхождение и качество мяса", "92% мяса / 90% мяса"),
    ("5294208319", "НЯМТОРИЯ Мясное разнообразие", "Собаки", "Лакомства", "PREMIUM", "Премиальное позиционирование", "мясо премиум-качества"),

    ("5294159811", "НЯМТОРИЯ Деликатесы из утки", "Собаки", "Лакомства", "NAT", "Натуральность состава", "Без глицерина; натуральное мясо премиум-качества"),
    ("5294159811", "НЯМТОРИЯ Деликатесы из утки", "Собаки", "Лакомства", "MEAT", "Происхождение и качество мяса", "94% мяса / 100% утка / 90% мяса"),
    ("5294159811", "НЯМТОРИЯ Деликатесы из утки", "Собаки", "Лакомства", "PREMIUM", "Премиальное позиционирование", "мясо премиум-качества"),

    ("5294111572", "НЯМТОРИЯ Вкусняшки мини 2.0", "Собаки", "Лакомства", "NAT", "Натуральность состава", "Без глицерина; натуральное мясо премиум-качества"),
    ("5294111572", "НЯМТОРИЯ Вкусняшки мини 2.0", "Собаки", "Лакомства", "MEAT", "Происхождение и качество мяса", "88% мяса / 90% мяса / 90% мяса"),
    ("5294111572", "НЯМТОРИЯ Вкусняшки мини 2.0", "Собаки", "Лакомства", "PREMIUM", "Премиальное позиционирование", "мясо премиум-качества"),

    ("5294843939", "Gosbi Cat Kitten", "Кошки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Mediterranean gourmet; Made in Spain"),
    ("5294960856", "Gosbi Cat Sterilised", "Кошки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Mediterranean gourmet; Made in Spain"),

    # ---- Доразбор 03.08 ----
    ("5306470815", "LiveRa Adult Fish", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% natural"),
    ("5306470815", "LiveRa Adult Fish", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic"),
    ("5306470815", "LiveRa Adult Fish", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "51% мяса"),

    ("5306440073", "LiveRa Adult ассорти из птицы (кошки)", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% natural"),
    ("5306440073", "LiveRa Adult ассорти из птицы (кошки)", "Кошки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic"),
    ("5306440073", "LiveRa Adult ассорти из птицы (кошки)", "Кошки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
    ("5306440073", "LiveRa Adult ассорти из птицы (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "54% мяса"),
    ("5306440073", "LiveRa Adult ассорти из птицы (кошки)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Супер премиум класс"),

    ("5306424879", "LiveRa Adult Large кролик (кошки)", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% natural"),
    ("5306424879", "LiveRa Adult Large кролик (кошки)", "Кошки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic"),
    ("5306424879", "LiveRa Adult Large кролик (кошки)", "Кошки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
    ("5306424879", "LiveRa Adult Large кролик (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "63% мяса"),
    ("5306424879", "LiveRa Adult Large кролик (кошки)", "Кошки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Корма из России"),
    ("5306424879", "LiveRa Adult Large кролик (кошки)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "супер премиум класса"),

    ("5308822504", "ХАУНДХОЛЛ лопаточный хрящ", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% Натурально"),

    ("5324402513", "KingKorm рубец бараний", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натурально; без красителей, без ГМО"),
    ("5324402513", "KingKorm рубец бараний", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Поддерживает здоровье зубов"),
    ("5324402513", "KingKorm рубец бараний", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Улучшает пищеварение"),
    ("5324402513", "KingKorm рубец бараний", "Собаки", "Лакомства", "MEAT", "Происхождение и качество мяса", "100% баранина"),

    ("5328154887", "All Cats Sterilized (AllerPetfood)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Complete premium pet food"),
    ("5328154887", "All Cats Sterilized (AllerPetfood)", "Кошки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Датское качество; AllerPetfood since 1921"),
    ("5328154887", "All Cats Sterilized (AllerPetfood)", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Omega-3, Omega-6; Contains Taurine"),
    ("5328154887", "All Cats Sterilized (AllerPetfood)", "Кошки", "Сухой корм", "DENT", "Стоматологическая польза", "Dental Care"),

    ("5325606109", "MEALGRAND индейка/курица (кошки)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Супер Премиум"),
    ("5325020749", "MEALGRAND ягнёнок (кошки)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Супер Премиум"),

    ("5254153857", "GRANDORF Cat курица", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free formula"),
    ("5254153857", "GRANDORF Cat курица", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Natural & Healthy"),
    ("5254153857", "GRANDORF Cat курица", "Кошки", "Влажный корм", "HOLISTIC", "Холистик-позиционирование", "Holistic Nutrition for Cats"),

    ("5254153733", "GRANDORF Cat тунец+креветки", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free formula"),
    ("5254153733", "GRANDORF Cat тунец+креветки", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Natural & Healthy"),
    ("5254153733", "GRANDORF Cat тунец+креветки", "Кошки", "Влажный корм", "HOLISTIC", "Холистик-позиционирование", "Holistic Nutrition for Cats"),

    ("5254153795", "GRANDORF Cat тунец+краб", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free formula"),
    ("5254153795", "GRANDORF Cat тунец+краб", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Natural & Healthy"),
    ("5254153795", "GRANDORF Cat тунец+краб", "Кошки", "Влажный корм", "HOLISTIC", "Холистик-позиционирование", "Holistic Nutrition for Cats"),

    ("5327393201", "Ваш Стандарт мясное ассорти", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "Без пафоса и переплат"),
    ("5327393201", "Ваш Стандарт мясное ассорти", "Собаки", "Сухой корм", "ORIGIN", "Страна производства / премиум-позиционирование", "Сделано на Кубани"),
    ("5327393201", "Ваш Стандарт мясное ассорти", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "ГОСТ"),

    ("5327553656", "Farmina VetLife Gastrointestinal", "Собаки", "Влажный корм", "VET", "Ветеринарная / лечебная диета", "Полнорационный диетический корм; здоровое ЖКТ"),
    ("5327553656", "Farmina VetLife Gastrointestinal", "Собаки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Поддержка здорового пищеварения; сильный иммунитет"),

    ("5259703870", "AWARD Sterilized кусочки в соусе (кошки)", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "100% BIO"),
    ("5259703870", "AWARD Sterilized кусочки в соусе (кошки)", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Urinary Health; Nutraceutical Formula"),
    ("5259703870", "AWARD Sterilized кусочки в соусе (кошки)", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Superpremium Cat Food"),
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
     "Отсутствие искусственных добавок, ГМО, «чистая» этикетка (без отдельного заявления Holistic — см. код HOLISTIC)",
     "100% натурально; без ГМО, красителей и ароматизаторов; 100% BIO; без глицерина"),
    ("HOLISTIC", "Холистик-позиционирование",
     "Отдельный маркетинговый термин «holistic» — философия рецептуры, часто идёт рядом с NAT/HYPO, но заявляется как отдельный статус, а не просто «натурально»",
     "Holistic; Holistic Nutrition; Holistic & Hypoallergenic"),
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
    ("ORIGIN", "Страна / регион производства",
     "Указание конкретной страны, региона или национальной принадлежности бренда (без оценки премиальности самой по себе)",
     "Made in Europe; Сделано в Таиланде; Сделано на Кубани; Произведено на Алтае; Mediterranean gourmet; «Корма из России»"),
    ("PREMIUM", "Премиальное позиционирование",
     "Заявление о премиум/супер-премиум статусе, бренд-наследии или знаках качества — независимо от страны производства",
     "Premium Quality; Superpremium; Professional; ГОСТ; «с 1965 года»"),
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
# Протеин / ингредиенты-клеймы / регион завода — нормализованные теги
# (для товара может быть несколько тегов протеина/ингредиентов; региона
# завода — не более одного). "—" в тегах не бывает: там, где на фасаде
# не указано и не применимо, товар просто отсутствует в списке тегов.
# =====================================================================
PROTEIN_TAGS = {
    "5214515666": ["Утка"], "5277005073": ["Ягнёнок"], "5276044539": ["Утка"],
    "5259070728": ["Индейка"], "5280647420": ["Курица"], "4685211777": ["Говядина"],
    "5110846419": ["Говядина"], "5277005334": ["Ягнёнок", "Индейка"], "5243322000": ["Индейка"],
    "5248949561": ["Тунец", "Лосось"], "4098300679": [],  # не мясной продукт
    "5293611532": ["Лосось", "Ягнёнок"], "5294915094": ["Курица", "Индейка"],
    "5294272736": ["Утка", "Тунец", "Курица"], "5293523232": ["Говядина"],
    "5306421483": ["Ягнёнок"], "5306617135": ["Индейка"], "5327569955": ["Утка"],
    "5327539554": [], "5327528788": [],  # не указан на фасаде
    "5327751575": ["Утка"], "5321602313": ["Говядина"], "5294994812": ["Белая рыба", "Лосось"],
    "5306618148": ["Индейка"], "5259705799": ["Индейка"], "5254153891": ["Курица", "Лосось"],
    "5322942235": [],  # "мясо" без конкретики

    # ---- Доразбор 01.08 + 03.08 (полный разбор, без "протеин-онли" сокращений) ----
    "5294201871": ["Курица", "Рыба"], "5294241487": ["Курица", "Рыба"],
    "5294353414": ["Ягнёнок"],
    "5294317994": ["Ягнёнок", "Утка"], "5294208319": ["Говядина", "Ягнёнок", "Курица"],
    "5294159811": ["Утка"], "5294111572": ["Говядина", "Утка", "Курица"],
    "5294843939": ["Курица", "Индейка"], "5294960856": ["Курица", "Лосось"],
    "5306470815": ["Белая рыба"], "5306440073": ["Птица (ассорти)"], "5306424879": ["Кролик"],
    "5308822504": [], "5324402513": ["Баранина"], "5321600961": [],
    "5328154887": ["Индейка"], "5325606109": ["Индейка", "Курица"], "5325020749": ["Ягнёнок"],
    "5254153857": ["Курица"], "5254153733": ["Тунец", "Креветки"], "5254153795": ["Тунец", "Краб"],
    "5327393201": [], "5327553656": [], "5259703870": ["Индейка"],
}

# Пары SKU одной линейки, отличающиеся только весом/размером упаковки — по
# просьбе пользователя такие дубли НЕ считаются повторно ни по одному признаку
# (клеймы/протеин/ингредиенты/завод уже учтены у представителя линейки):
# 5306446981, 5306427606 (дубль 5306421483), 5327563858 (дубль 5327569955),
# 5327549511 (дубль 5327539554), 5295223041 (дубль 5277005073),
# 5321605313 (дубль 5321602313), 5306515614 (дубль 5306618148).
EXTRA_PROTEIN_ONLY = {}

PROTEIN_UNKNOWN_OIDS = {
    "4098300679", "5327539554", "5327528788", "5322942235",
    "5308822504", "5321600961", "5327393201", "5327553656",
}
PROTEIN_TOTAL_SKUS = len(PROTEIN_TAGS) + len(EXTRA_PROTEIN_ONLY)

INGREDIENT_TAGS = {
    "5277005073": ["Батат", "Омега-3/6", "Глюкозамин/хондроитин", "Пробиотики"],
    "5276044539": ["Яблоко", "Артишок"],
    "4098300679": ["Овёс", "Витамины", "Фолиевая кислота"],
    "5294272736": ["Морская капуста (водоросли)", "Морковь"],
    "5306421483": ["Рис"], "5306617135": ["Рис"], "5327569955": ["Картофель"],
    "5306618148": ["Рис"], "5259705799": ["Овощи"],
    "5294353414": ["Рис"], "5294208319": ["Морская капуста (водоросли)"],
    "5294159811": ["Морская капуста (водоросли)"], "5294111572": ["Морская капуста (водоросли)"],
    "5306470815": ["Рис"], "5306424879": ["Рис"], "5328154887": ["Омега-3/6"],
}

# Правило: для России — название юрлица/завода (если его не видно на этикетке —
# считаем неизвестным, одного региона/города недостаточно). Для остальных стран —
# достаточно названия страны.
PLANT_REGION = {
    "5214515666": None, "5277005073": "Бельгия",
    "5276044539": "ООО «ПетЭксперт» (Россия)",
    "5259070728": None, "5280647420": None,  # Липецкая обл. известна, юрлицо не подтверждено
    "4685211777": None,
    "5110846419": None,  # Волгоград известен, ИП не читается
    "5277005334": None,  # только "Европа", страна не уточнена
    "5243322000": "ООО «ПроРацион» (Россия)", "5248949561": "Таиланд", "4098300679": None,
    "5293611532": None, "5294915094": None,  # Gosbi — Испания не подтверждена по этикетке
    "5294272736": None, "5293523232": None,
    "5306421483": None, "5306617135": None,  # "Корма из России" — страна известна, юрлицо нет
    "5327569955": "Италия", "5327539554": "Италия", "5327528788": "Италия",
    "5327751575": "ИП Клюкова А.А. (Россия)",
    "5321602313": "Китай",  # для не-России юрлицо не нужно, достаточно страны
    "5294994812": None,
    "5306618148": None, "5259705799": None, "5254153891": None,
    "5322942235": None,  # Алтайский край известен, юрлицо нет
    "5294201871": "Испания", "5294241487": "Испания", "5294353414": "Испания",
    "5294843939": "Испания", "5294960856": "Испания",
    "5294317994": None, "5294208319": None, "5294159811": None, "5294111572": None,
    "5306470815": None, "5306440073": None, "5306424879": None,
    "5308822504": None, "5324402513": None, "5321600961": None,
    "5328154887": None,  # Дания — страна бренда AllerPetfood, не завод-изготовитель
    "5325606109": None, "5325020749": None,
    "5254153857": None, "5254153733": None, "5254153795": None,
    "5327393201": "ИП Клюкова А.А. (Россия)",
    "5327553656": "Италия",
    "5259703870": None,
}

ws_pt = wb.create_sheet("Протеин-Ингредиенты-Завод")
headers_pt = ["ID товара", "Название товара", "Тип тега", "Значение"]
for i, h in enumerate(headers_pt, start=1):
    ws_pt.cell(row=1, column=i, value=h)
style_header_row(ws_pt, 1, len(headers_pt))

name_by_oid = {p["oid"]: p["name"] for p in products}
name_by_oid.update({oid: name for oid, (name, _) in EXTRA_PROTEIN_ONLY.items()})

pt_rows = []
for oid, tags in PROTEIN_TAGS.items():
    for tag in tags:
        pt_rows.append((oid, name_by_oid[oid], "Протеин", tag))
for oid, (_, tags) in EXTRA_PROTEIN_ONLY.items():
    for tag in tags:
        pt_rows.append((oid, name_by_oid[oid], "Протеин", tag))
for oid, tags in INGREDIENT_TAGS.items():
    for tag in tags:
        pt_rows.append((oid, name_by_oid[oid], "Ингредиент-клеймо", tag))
for oid, region in PLANT_REGION.items():
    pt_rows.append((oid, name_by_oid[oid], "Производитель", region if region else "Неизвестно"))

for i, row in enumerate(pt_rows, start=2):
    for j, val in enumerate(row, start=1):
        c = ws_pt.cell(row=i, column=j, value=val)
        c.font = base_font
        c.alignment = wrap
        c.border = border
autosize(ws_pt, [12, 34, 16, 30])
ws_pt.freeze_panes = "A2"

# =====================================================================
# Sheet 4: Аналитика
# =====================================================================
ws4 = wb.create_sheet("Аналитика")
ws4["A1"] = "Аналитика: популярность направлений клеймов среди новинок"
ws4["A1"].font = title_font
ws4["A2"] = (f"Выборка: {len(products)} карточек (собаки+кошки, корма и лакомства) за 3 снимка — "
             "11 отобраны 2026-07-31 через сортировку «Новинки» (не подтверждено ID), "
             "13 добавлены 2026-08-01 и 27 добавлены 2026-08-03 — все 40 подтверждены сравнением ID с базой "
             "(объективно новые). Полностью разобраны (клеймы+протеин+ингредиенты+завод) все объективно новые SKU, "
             "кроме 7 дублей по весу/размеру упаковки той же рецептуры (не считаются повторно ни по одному признаку).")
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

# =====================================================================
# Доп. аналитика: Протеин / Ингредиенты-клеймы / Регион завода
# Базой 100% для протеина и региона завода служит количество товаров,
# где значение удалось определить (неизвестные исключены из знаменателя,
# как и для завода-изготовителя в целом). Для ингредиентов-клеймов базой
# остаётся вся выборка — отсутствие такого клейма само по себе значимо.
# =====================================================================
protein_known = PROTEIN_TOTAL_SKUS - len(PROTEIN_UNKNOWN_OIDS)
region_known = sum(1 for v in PLANT_REGION.values() if v)

def add_tag_table(ws, start_row, title, tag_type, denominator, denom_label, exclude=()):
    ws.cell(row=start_row, column=1, value=title).font = Font(name=FONT_NAME, bold=True, size=12)
    hrow_t = start_row + 2
    for i, h in enumerate(["Значение", "Товаров", "Доля"], start=1):
        ws.cell(row=hrow_t, column=i, value=h)
    style_header_row(ws, hrow_t, 3)
    tags = sorted({row[3] for row in pt_rows if row[2] == tag_type and row[3] not in exclude})
    r0 = hrow_t + 1
    for i, tag in enumerate(tags):
        rr = r0 + i
        ws.cell(row=rr, column=1, value=tag)
        ws.cell(row=rr, column=2,
                value=f'=COUNTIFS(\'Протеин-Ингредиенты-Завод\'!C:C,"{tag_type}",\'Протеин-Ингредиенты-Завод\'!D:D,A{rr})')
        ws.cell(row=rr, column=3, value=f'=B{rr}/{denominator}')
        ws.cell(row=rr, column=3).number_format = "0%"
        for c in range(1, 4):
            ws.cell(row=rr, column=c).font = base_font
            ws.cell(row=rr, column=c).border = border
            ws.cell(row=rr, column=c).alignment = wrap
    last_r = r0 + len(tags) - 1
    note_r = last_r + 1
    ws.cell(row=note_r, column=1, value=f"База 100% = {denominator} ({denom_label})").font = note_font
    ws.merge_cells(f"A{note_r}:D{note_r}")
    return note_r

ws4.column_dimensions["A"].width = 34
next_row = last_data_row + 4
next_row = add_tag_table(ws4, next_row,
                          f"Протеин — популярность по ВСЕМ SKU новых линеек (не по одному товару на линейку), {PROTEIN_TOTAL_SKUS} SKU всего",
                          "Протеин", protein_known,
                          f"{protein_known} SKU, где протеин назван на фасаде, из {PROTEIN_TOTAL_SKUS} SKU всего "
                          f"(15 подробно разобранных + 34 объективно новых за 2026-08-03, включая непоказанные фото/клеймами вариации линеек)")
next_row = add_tag_table(ws4, next_row + 3, "Ингредиенты-клеймы (фасад) — популярность среди новинок",
                          "Ингредиент-клеймо", n_products,
                          f"вся выборка ({n_products}) — отсутствие клейма тоже значимо")
next_row = add_tag_table(ws4, next_row + 3,
                          "Производитель — название завода для России, страна для остальных (среди идентифицированных)",
                          "Производитель", region_known,
                          f"{region_known} товаров, где производитель однозначно определён (для РФ — по названию юрлица, "
                          f"иначе — не хватает данных и товар не учитывается), из {n_products} в подробно разобранной выборке",
                          exclude={"Неизвестно"})
analytics_end_row = next_row

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
concl_row = max(analytics_end_row, hrow4 + 24) + 4
ws4.cell(row=concl_row, column=1, value="Выводы (наблюдения: 2026-07-31, 2026-08-01, 2026-08-03)").font = Font(name=FONT_NAME, bold=True, size=12)
conclusions = [
    "1. Натуральность состава (NAT) — самое частое направление даже без учёта Holistic: 29 из 51 карточки (57%) "
    "заявляют «100% натурально», «без ГМО» или «без глицерина». Это «входной билет» для нового SKU почти в любом "
    "сегменте и ценовой категории, а не способ выделиться.",
    "2. Холистик (HOLISTIC) — отдельный от «натуральности» маркетинговый статус рецептуры: встречается у 12 из 51 "
    "(24%) — почти исключительно у GRANDORF и LiveRa, и всегда вместе с HYPO или NAT, никогда сам по себе — то есть "
    "используется как усилитель, а не самостоятельный аргумент.",
    "3. Ветеринарная/лечебная диета (VET) — 5 из 51 (10%), почти целиком за счёт линейки Farmina VetLife "
    "(Hypoallergenic, Renal, Diabetic, Gastrointestinal) — признак того, что лечебное питание становится "
    "самостоятельным растущим направлением, а не нишей одного бренда.",
    "4. Завод-изготовитель вскрыл географический кластер контрактного производства в России: ООО «ПетЭксперт» "
    "(Липецкая обл., Елецкий р-н, ОЭЗ «Липецк») делает корм под брендом AWARD для АО «Валта Пет Продактс» и, "
    "вероятно, тот же кластер — под CLAN VET; ООО «ПроРацион» (Новосибирск) делает EDWIN; бренд LiveRa указывает "
    "«Корма из России» с адресом в Барнауле. Как и с GRANDORF/United Petfood в Европе, один и тот же завод может "
    "стоять за несколькими разными брендами на полке Ozon — бренд и реальный изготовитель систематически различны.",
    "5. Протеиновый разрыв премиум/эконом: индейка и утка — основные протеины у гипоаллергенных и "
    "премиальных SKU (EDWIN, LiveRa Hypoallergenic, Farmina Duck, Premier, GRANDORF) — как менее аллергенные и "
    "более «премиальные» виды мяса; говядина доминирует в эконом-сегменте и базовых лакомствах (Мурчик, "
    "DelicaChoice, Лопай лапой, Альма) — как самое доступное и привычное мясное сырьё.",
    "6. Сегментный разрыв в клеймах подтверждён на трёх независимых эконом-товарах: «Мурчик» (влажный "
    "корм для кошек), «DelicaChoice» (Китай) и «DOGNIS» — оба последних прессованные кости из жил, все три "
    "объективно новые SKU, и ни один не несёт ни одного клейма из справочника, только вес/размер/базовое описание. "
    "Messaging-борьба почти полностью сосредоточена в премиум- и средне-премиальном сегменте.",
    "7. Важный нюанс по составу: клейм не всегда универсален внутри одной товарной линейки. У НЯМТОРИЯ «без "
    "глицерина» подтверждается для вкусов утка и курица, но вкус тунец, по составу на этикетке, глицерин "
    "СОДЕРЖИТ — «чистая этикетка» может относиться только к части ассортимента бренда, а не ко всей линейке.",
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
