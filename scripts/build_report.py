\
# -*- coding: utf-8 -*-
"""
Строит отчёт ozon-claims-tracker/reports/claims_report.xlsx
Пилотный запуск (Day 0): примеры отобраны через сортировку "Новинки" на Ozon
как удобный (но не полностью надёжный) источник кандидатов.
Начиная со следующего запуска, реальные новые карточки определяются сравнением
с базой data/products_history.csv (по ID товара).
"""
import io
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
    # Доразбор 04.08 — новый снимок, объективно новые SKU (ID выше макс. на 03.08)
    "5345830359": ("2026-08-04", "5345830359_petsbrunch.jpg"),
    "5358300410": ("2026-08-04", "5358300410_generic10kg.jpg"),
    "5357288481": ("2026-08-04", "5357288481_farmina_quail.jpg"),
    "5358242931": ("2026-08-04", "5358242931_rodnye_skoblyanka.jpg"),
    "5358240926": ("2026-08-04", "5358240926_rodnye_baranina.jpg"),
    "5359892602": ("2026-08-04", "5359892602_puffins_turkeysalmon.jpg"),
    "5359453152": ("2026-08-04", "5359453152_puffins_beefsalmon.jpg"),
    "5359660535": ("2026-08-04", "5359660535_puffins_assorti.jpg"),
    "5359822345": ("2026-08-04", "5359822345_grandorf_lambturkey.jpg"),
    "5357563517": ("2026-08-04", "5357563517_delish_chicken.jpg"),
    "5357853793": ("2026-08-04", "5357853793_delish_beef.jpg"),
    "5331257302": ("2026-08-04", "5331257302_pauchi_turkeycranberry.jpg"),
    "5358325770": ("2026-08-04", "5358325770_kremlakomstvo.jpg"),
    "5358242441": ("2026-08-04", "5358242441_mnyams_tuna.jpg"),
    "5358240663": ("2026-08-04", "5358240663_mnyams_beef.jpg"),
    "5358241683": ("2026-08-04", "5358241683_mnyams_chicken.jpg"),
    "5358240549": ("2026-08-04", "5358240549_4hvostika_trava.jpg"),
    # Day 6 (2026-08-12) — подтверждено сравнением ID с базой (ID выше макс. на 2026-08-04)
    "5430117217": ("2026-08-12", "5430117217_kasha_beef.jpg"),
    "5423933879": ("2026-08-12", "5423933879_kasha.jpg"),
    "5423741999": ("2026-08-12", "5423741999_kasha.jpg"),
    "5432594014": ("2026-08-12", "5432594014_kasha_fish.jpg"),
    "5419467431": ("2026-08-12", "5419467431_kasha.jpg"),
    "5418435027": ("2026-08-12", "5418435027_kasha.jpg"),
    "5417460160": ("2026-08-12", "5417460160_alma.jpg"),
    "5432515655": ("2026-08-12", "5432515655_cat_2.jpg"),
    "5431773123": ("2026-08-12", "5431773123_mnyams.jpg"),
    "5384510562": ("2026-08-12", "5384510562_blitz.jpg"),
    "5429960930": ("2026-08-12", "5429960930_animonda.jpg"),
    "5430728946": ("2026-08-12", "5430728946_clan_vet.jpg"),
    "5430727673": ("2026-08-12", "5430727673_royalcanin.jpg"),
    "5430194418": ("2026-08-12", "5430194418_animonda_vf.jpg"),
    "5424780704": ("2026-08-12", "5424780704_farmina_ocean_dog.jpg"),
    "5424781094": ("2026-08-12", "5424781094_farmina_quinoa_lamb.jpg"),
    "5424780086": ("2026-08-12", "5424780086_farmina_ancestral_chicken.jpg"),
    "5424781029": ("2026-08-12", "5424781029_farmina_ancestral_chicken_ster.jpg"),
    "5424783286": ("2026-08-12", "5424783286_farmina_ancestral_lamb.jpg"),
    "5424776640": ("2026-08-12", "5424776640_farmina_pumpkin_quail.jpg"),
    "5424785573": ("2026-08-12", "5424785573_farmina_pumpkin_boar.jpg"),
    "5424783287": ("2026-08-12", "5424783287_farmina_pumpkin_lamb_cat.jpg"),
    "5424784781": ("2026-08-12", "5424784781_farmina_quinoa_duck.jpg"),
    "5424782869": ("2026-08-12", "5424782869_farmina_prime_kitten.jpg"),
    "5428395063": ("2026-08-12", "5428395063_titbit_fresh.jpg"),
    "5428925488": ("2026-08-12", "5428925488_titbit_dent_beef.jpg"),
    "5429259395": ("2026-08-12", "5429259395_titbit_dent_tvorog.jpg"),
    "5431102190": ("2026-08-12", "5431102190_titbit_dent_veal.jpg"),
    "5430211436": ("2026-08-12", "5430211436_award_veal.jpg"),
    "5432604231": ("2026-08-12", "5432604231_karmy_starter.jpg"),
    "5432823007": ("2026-08-12", "5432823007_pedigree_dentastix_light.jpg"),
    # Доразбор 12.08 — пропущенные при первом проходе, найдены аудитом ?sorting=new
    "5433103013": ("2026-08-12", "5433103013_clan_vet_gastro.jpg"),
    "5433102619": ("2026-08-12", "5433102619_clan_classic_rabbit.jpg"),
    "5433103250": ("2026-08-12", "5433103250_berkley4.jpg"),
    "5433103558": ("2026-08-12", "5433103558_berkley6.jpg"),
    "5433104303": ("2026-08-12", "5433104303_berkley7.jpg"),
    "5433102991": ("2026-08-12", "5433102991_delish_rabbit.jpg"),
    "5433102865": ("2026-08-12", "5433102865_clan_pride.jpg"),
    "5433103171": ("2026-08-12", "5433103171_delish_beef_lamb.jpg"),
    "5436654518": ("2026-08-12", "5436654518_duck_heads.jpg"),
    "5424786354": ("2026-08-12", "5424786354_farmina_prime_dog.jpg"),
    "5424780417": ("2026-08-12", "5424780417_farmina_ancestral_dog_lamb.jpg"),
    # Day 7 (2026-08-24) — Каша для собак
    "5542503462": ("2026-08-24", "5542503462_check.jpg"),
    "5445105149": ("2026-08-24", "5445105149_alma_fish.jpg"),
    "5445127064": ("2026-08-24", "5445127064_alma_grechka.jpg"),
    "5456129061": ("2026-08-24", "5456129061_etizveri_lapazlak.jpg"),
    "5456129661": ("2026-08-24", "5456129661_etizveri_rice.jpg"),
    "5456131380": ("2026-08-24", "5456131380_etizveri.jpg"),
    "5456131513": ("2026-08-24", "5456131513_etizveri_grechka.jpg"),
    "5479813617": ("2026-08-24", "5479813617_generic10kg.jpg"),
    "5500647058": ("2026-08-24", "5500647058_sibkorm_assorti.jpg"),
    # Day 7 (2026-08-24) — Сухой корм для собак
    "5515836917": ("2026-08-24", "5515836917_wanpy_chicken.jpg"),
    "5515838796": ("2026-08-24", "5515838796_wanpy_salmon.jpg"),
    "5524575556": ("2026-08-24", "5524575556_wanpy_lamb.jpg"),
    "5529408722": ("2026-08-24", "5529408722_generic_turkey.jpg"),
    "5530070606": ("2026-08-24", "5530070606_provipet_lamb.jpg"),
    "5542850495": ("2026-08-24", "5542850495_award_salmon_new.jpg"),
    "5555438805": ("2026-08-24", "5555438805_provipet_turkey.jpg"),
    "5556742789": ("2026-08-24", "5556742789_alfapet_fish.jpg"),
    "5566163855": ("2026-08-24", "5566163855_grandorf_turkey.jpg"),
    "5566490497": ("2026-08-24", "5566490497_alphapet_beef.jpg"),
    "5568844559": ("2026-08-24", "5568844559_vivere_lamb.jpg"),
    "5574603661": ("2026-08-24", "5574603661_proplan_gastro.jpg"),
    "5578347424": ("2026-08-24", "5578347424_hilton_beef.jpg"),
    # Day 7 (2026-08-24) — Влажный корм для собак
    "5471075781": ("2026-08-24", "5471075781_generic_veal.jpg"),
    "5530369991": ("2026-08-24", "5530369991_wanpy_wet_chicken.jpg"),
    "5530650366": ("2026-08-24", "5530650366_wanpy_wet_beef.jpg"),
    "5530650986": ("2026-08-24", "5530650986_wanpy_wet_lamb.jpg"),
    "5544614194": ("2026-08-24", "5544614194_royalcanin_mini.jpg"),
    "5549300374": ("2026-08-24", "5549300374_pedigree_puppy.jpg"),
    "5556733970": ("2026-08-24", "5556733970_zoogurman_turkey.jpg"),
    "5556739255": ("2026-08-24", "5556739255_zoogurman_veal_rumen.jpg"),
    # Day 7 (2026-08-24) — Лакомства для собак
    "5544395456": ("2026-08-24", "5544395456_legkoe_govyazhe.jpg"),
    "5555142584": ("2026-08-24", "5555142584_biohaus_duck.jpg"),
    "5555380815": ("2026-08-24", "5555380815_luckers_apple_beef.jpg"),
    "5555663607": ("2026-08-24", "5555663607_biohaus_sushi_rabbit_cod.jpg"),
    "5556221688": ("2026-08-24", "5556221688_biohaus_sushi_turkey_cod.jpg"),
    "5556538078": ("2026-08-24", "5556538078_biohaus_shashlik_chicken.jpg"),
    "5556653006": ("2026-08-24", "5556653006_biohaus_ears_rabbit_duck.jpg"),
    "5577356638": ("2026-08-24", "5577356638_titbit_sushki.jpg"),
    "5577356979": ("2026-08-24", "5577356979_titbit_biscotti.jpg"),
    "5577357024": ("2026-08-24", "5577357024_titbit_dent_cheese.jpg"),
    "5577357304": ("2026-08-24", "5577357304_titbit_dent_duck.jpg"),
    "5577357359": ("2026-08-24", "5577357359_titbit_salmon_straw.jpg"),
    "5577358728": ("2026-08-24", "5577358728_titbit_salmon_cubes.jpg"),
    "5577359119": ("2026-08-24", "5577359119_titbit_cracker.jpg"),
    "5577402177": ("2026-08-24", "5577402177_beef_trachea.jpg"),
    "5585499483": ("2026-08-24", "5585499483_biohaus_sandwich_duck.jpg"),
    # Day 7 (2026-08-24) — Сухой корм для кошек
    "5515840508": ("2026-08-24", "5515840508_wanpy_cat_chicken.jpg"),
    "5515844102": ("2026-08-24", "5515844102_wanpy_cat_salmon.jpg"),
    "5542531112": ("2026-08-24", "5542531112_award_cat_salmon.jpg"),
    "5542594928": ("2026-08-24", "5542594928_generic_privered.jpg"),
    "5543934520": ("2026-08-24", "5543934520_alphapet_sensitive.jpg"),
    "5543935300": ("2026-08-24", "5543935300_goodwin_turkey_duck.jpg"),
    "5543935895": ("2026-08-24", "5543935895_alphapet_large.jpg"),
    "5543936157": ("2026-08-24", "5543936157_goodwin_kitten.jpg"),
    "5543936716": ("2026-08-24", "5543936716_alphapet_sterilized_chicken.jpg"),
    "5554107233": ("2026-08-24", "5554107233_primordial_duck_turkey.jpg"),
    "5554109198": ("2026-08-24", "5554109198_primordial_salmon_quail.jpg"),
    "5554109696": ("2026-08-24", "5554109696_primordial_salmon_turkey.jpg"),
    "5555801791": ("2026-08-24", "5555801791_tasty_beef.jpg"),
    "5556095810": ("2026-08-24", "5556095810_monge_duck.jpg"),
    "5556862888": ("2026-08-24", "5556862888_generic_sterilized.jpg"),
    "5566319732": ("2026-08-24", "5566319732_alphapet_sp_turkey_kitten.jpg"),
    "5566555324": ("2026-08-24", "5566555324_alphapet_senior_rabbit_duck.jpg"),
    "5566602205": ("2026-08-24", "5566602205_alphapet_beef_liver.jpg"),
    "5566632364": ("2026-08-24", "5566632364_alphapet_menu_beef_cat.jpg"),
    "5566657335": ("2026-08-24", "5566657335_alphapet_sp_chicken.jpg"),
    "5568745202": ("2026-08-24", "5568745202_karmy_cat_allergy.jpg"),
    "5568879450": ("2026-08-24", "5568879450_optimanova_rabbit.jpg"),
    "5574288140": ("2026-08-24", "5574288140_proplan_cat_gastro.jpg"),
    "5578339795": ("2026-08-24", "5578339795_hilton_cat_rabbit.jpg"),
    # Day 7 (2026-08-24) — Влажный корм для кошек
    "5484409205": ("2026-08-24", "5484409205_gourmet_gold_chicken.jpg"),
    "5484413359": ("2026-08-24", "5484413359_gourmet_gold_turkey.jpg"),
    "5484418520": ("2026-08-24", "5484418520_gourmet_gold_tuna.jpg"),
    "5530649635": ("2026-08-24", "5530649635_wanpy_cat_wet_tuna.jpg"),
    "5530649765": ("2026-08-24", "5530649765_wanpy_cat_wet_chicken_shrimp.jpg"),
    "5530650056": ("2026-08-24", "5530650056_wanpy_cat_wet_chicken_salmon.jpg"),
    "5530367150": ("2026-08-24", "5530367150_wanpy_cat_wet_chicken_scallop.jpg"),
    "5530650988": ("2026-08-24", "5530650988_wanpy_cat_wet_salmon_tuna_stew.jpg"),
    "5530651154": ("2026-08-24", "5530651154_wanpy_cat_wet_chicken_duck_stew.jpg"),
    "5530651642": ("2026-08-24", "5530651642_wanpy_cat_wet_chicken_cod.jpg"),
    "5530653845": ("2026-08-24", "5530653845_wanpy_cat_wet_chicken_tuna.jpg"),
    "5530654016": ("2026-08-24", "5530654016_wanpy_cat_wet_chicken_crab.jpg"),
    "5538615147": ("2026-08-24", "5538615147_generic_turkey_chicken.jpg"),
    "5543025974": ("2026-08-24", "5543025974_mnyams_delicates.jpg"),
    "5544350676": ("2026-08-24", "5544350676_generic_beef.jpg"),
    "5544357753": ("2026-08-24", "5544357753_generic_meat_assorti.jpg"),
    "5551551944": ("2026-08-24", "5551551944_royalcanin_kitten.jpg"),
    "5566574938": ("2026-08-24", "5566574938_schesir_chicken_egg.jpg"),
    "5566598587": ("2026-08-24", "5566598587_schesir_chicken.jpg"),
    "5566620215": ("2026-08-24", "5566620215_schesir_chicken_beef.jpg"),
    "5566642607": ("2026-08-24", "5566642607_schesir_chicken_duck.jpg"),
    # Day 7 (2026-08-24) — Лакомства для кошек
    "5542724663": ("2026-08-24", "5542724663_hearts_turkey_cream.jpg"),
    "5543308856": ("2026-08-24", "5543308856_ferma_kota_sticks_rabbit.jpg"),
    "5543534231": ("2026-08-24", "5543534231_ferma_kota_sticks_salmon.jpg"),
    "5543551335": ("2026-08-24", "5543551335_ferma_kota_sticks_duck.jpg"),
    "5549365822": ("2026-08-24", "5549365822_dreamies_milkshake.jpg"),
    "5579052755": ("2026-08-24", "5579052755_award_special_care_duck_tuna.jpg"),
    "5579058668": ("2026-08-24", "5579058668_award_special_care_chicken.jpg"),
    "5579066625": ("2026-08-24", "5579066625_award_special_care_salmon_tuna.jpg"),
    "5582640835": ("2026-08-24", "5582640835_sticks_chicken.jpg"),
    "5584169750": ("2026-08-24", "5584169750_murkel_chicken.jpg"),
    "5584218651": ("2026-08-24", "5584218651_murkel_flax_chicken_liver.jpg"),
    "5584264918": ("2026-08-24", "5584264918_murkel_tuna.jpg"),
    # Day 8 (2026-09-02) — Сухой корм для кошек
    "5658764220": ("2026-09-02", "5658764220_rodnyekorma_cat.jpg"),
    "5670474832": ("2026-09-02", "5670474832_award_hairball.jpg"),
    "5670477242": ("2026-09-02", "5670477242_award_urinary.jpg"),
    "5670498757": ("2026-09-02", "5670498757_award_indoor_big.jpg"),
    "5670499419": ("2026-09-02", "5670499419_award_sterilized_turkey.jpg"),
    "5670501608": ("2026-09-02", "5670501608_award_sterilized_fish.jpg"),
    # Day 8 (2026-09-02) — Влажный корм для собак
    "5669273502": ("2026-09-02", "5669273502_yummy_turkey.jpg"),
    "5670094001": ("2026-09-02", "5668919115_check.jpg"),
    "5670712472": ("2026-09-02", "5670714887_mnyams_ferma_fish.jpg"),
    "5670714813": ("2026-09-02", "5671464738_mnyams_duck_mousse.jpg"),
    "5670718836": ("2026-09-02", "5671464738_mnyams_duck_mousse.jpg"),
    "5670720625": ("2026-09-02", "5670720986_mnyams_rubets.jpg"),
    "5671464738": ("2026-09-02", "5671464738_mnyams_duck_mousse.jpg"),
    "5671505639": ("2026-09-02", "5669273502_yummy_turkey.jpg"),
    # Day 8 (2026-09-02) — Сухой корм для собак
    "5623144132": ("2026-09-02", "5623144132_farmina_prime_dog_lamb.jpg"),
    "5634101309": ("2026-09-02", "5634101309_grandorf_vet_hypo.jpg"),
    "5662623115": ("2026-09-02", "5662623115_grandorf_fresh_salmon.jpg"),
    "5662623376": ("2026-09-02", "5662623376_grandorf_fresh_turkey.jpg"),
    "5662626493": ("2026-09-02", "5662626493_grandorf_fresh_duck.jpg"),
    "5670255238": ("2026-09-02", "5670255238_royalcanin_neutered_mini.jpg"),
    "5670469697": ("2026-09-02", "5670469697_award_hypo_fish.jpg"),
    # Day 8 (2026-09-02) — Влажный корм для кошек
    "5654017462": ("2026-09-02", "5654017462_alphapet_wow_rabbit.jpg"),
    "5664781811": ("2026-09-02", "5664781811_peppo_chicken.jpg"),
    "5664784502": ("2026-09-02", "5664784502_peppo_lamb.jpg"),
    "5668947557": ("2026-09-02", "5668947557_generic_beef_cat.jpg"),
    "5669095176": ("2026-09-02", "5669095176_vkusnaya_miska_rabbit.jpg"),
    "5670228940": ("2026-09-02", "5670228940_whiskas.jpg"),
    "5670237889": ("2026-09-02", "5670237889_sheba_rabbit_turkey.jpg"),
    "5670251820": ("2026-09-02", "5670251820_sheba_chicken.jpg"),
    "5670328613": ("2026-09-02", "5670328613_perfectfit_turkey.jpg"),
    "5670392893": ("2026-09-02", "5670392893_royalcanin_sterilised_gravy.jpg"),
    "5670714887": ("2026-09-02", "5670714887_mnyams_ferma_fish.jpg"),
    "5670716904": ("2026-09-02", "5670716904_mnyams_ferma_turkey.jpg"),
    "5670719609": ("2026-09-02", "5670719609_mnyams_ferma_veal.jpg"),
    "5670721156": ("2026-09-02", "5670721156_mnyams_ferma_beef.jpg"),
    "5671557401": ("2026-09-02", "5671557401_mnyams_duck.jpg"),
    # Day 8 (2026-09-02) — Лакомства для собак
    "5632614860": ("2026-09-02", "5632614860_inaba_neji.jpg"),
    "5632615231": ("2026-09-02", "5632615231_inaba_churubee.jpg"),
    "5632624539": ("2026-09-02", "5632624539_inaba_churu_skincoat.jpg"),
    "5654589093": ("2026-09-02", "5654589093_inaba_churuball_veg.jpg"),
    "5654591896": ("2026-09-02", "5654591896_inaba_churu_gi.jpg"),
    "5654592771": ("2026-09-02", "5654592771_inaba_churuball.jpg"),
    "5654599026": ("2026-09-02", "5654599026_inaba_tobitsuku_skin.jpg"),
    "5654599608": ("2026-09-02", "5654599608_inaba_tobitsuku_gi.jpg"),
    "5670711483": ("2026-09-02", "5670711483_mnyams_dental_salmon.jpg"),
    "5670715414": ("2026-09-02", "5670715414_mnyams_kolbaski.jpg"),
    "5670717416": ("2026-09-02", "5670717416_mnyams_chicken_breast.jpg"),
    "5670718664": ("2026-09-02", "5670718664_mnyams_filemignon.jpg"),
    "5670720986": ("2026-09-02", "5670720986_mnyams_rubets.jpg"),
    "5671459555": ("2026-09-02", "5671459555_mnyams_dental_beef.jpg"),
    "5671490575": ("2026-09-02", "5671490575_mnyams_dental_bones.jpg"),
    "5671509904": ("2026-09-02", "5671509904_mnyams_trachea.jpg"),
    "5671561038": ("2026-09-02", "5671561038_mnyams_lamb_sausages.jpg"),
    # Day 8 (2026-09-02) — Лакомства для кошек
    "5632614370": ("2026-09-02", "5632614370_inaba_ciaopaste.jpg"),
    "5632614943": ("2026-09-02", "5632614943_inaba_churu_tuna_katsuo.jpg"),
    "5632615075": ("2026-09-02", "5632615075_inaba_yakisasami_crab.jpg"),
    "5632615106": ("2026-09-02", "5632615106_inaba_yakimiksu.jpg"),
    "5632619730": ("2026-09-02", "5632619730_inaba_churu_cat_skincoat.jpg"),
    "5654586643": ("2026-09-02", "5654586643_inaba_churu_tuna_maguro.jpg"),
    "5654587557": ("2026-09-02", "5654587557_inaba_tobitsuku_cat_broth.jpg"),
    "5654590710": ("2026-09-02", "5654590710_inaba_churu_ph.jpg"),
    "5654591217": ("2026-09-02", "5654591217_inaba_ponchuru_tuna.jpg"),
    "5654591653": ("2026-09-02", "5654591653_inaba_churu_chicken_crab.jpg"),
    "5654592626": ("2026-09-02", "5654592626_inaba_churu_tuna_scallop.jpg"),
    "5654594462": ("2026-09-02", "5654594462_inaba_churu_tuna_lobster.jpg"),
    "5654596008": ("2026-09-02", "5654596008_inaba_churu_urinary.jpg"),
    "5654598009": ("2026-09-02", "5654598009_inaba_churu_chicken_squid.jpg"),
    "5654598442": ("2026-09-02", "5654598442_inaba_tobitsuku_scallop.jpg"),
    "5654598890": ("2026-09-02", "5654598890_inaba_churu_kitten.jpg"),
    "5654599052": ("2026-09-02", "5654599052_inaba_yakisasami_beef.jpg"),
    "5654599091": ("2026-09-02", "5654599091_inaba_ponchuru_chicken_scallop.jpg"),
    "5654599124": ("2026-09-02", "5654599124_inaba_churu_chicken_shrimp.jpg"),
    "5670711038": ("2026-09-02", "5670711038_mnyams_crem_duck_tuna.jpg"),
    "5671503941": ("2026-09-02", "5671503941_mnyams_crem_chicken.jpg"),
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

ws["A1"] = "Клеймы новых карточек — Корма и лакомства для собак и кошек (Ozon)"
ws["A1"].font = title_font
ws["A2"] = ("2026-07-31 (Day 0): кандидаты отобраны через фильтр Ozon «Новинки» — быстрый, но не полностью "
            "надёжный источник (Ozon продвигает не только объективно новые товары). "
            "Начиная с 2026-08-01: товары подтверждены объективно — сравнением ID с базой "
            "data/products_history.csv; их ID выше максимума, зафиксированного в предыдущем снимке той же категории. "
            "ДВА ОГРАНИЧЕНИЯ МЕТОДА: (1) Ложные срабатывания — «объективно новый ID» доказывает лишь то, что "
            "карточка на Ozon появилась впервые в поле зрения, не то, что сам товар впервые вышел на рынок. "
            "Из-за высокой доли 3P-продаж на Ozon один и тот же товар может получить новую отдельную карточку с "
            "новым ID при листинге ещё одним продавцом (например, 20-м по счёту) или при повторном листинге тем же "
            "продавцом — метод не всегда отличает это от действительно нового товара; распознанные вручную дубли "
            "(по фото/описанию) не учитываются повторно в аналитике, но не все подобные случаи гарантированно "
            "пойманы. (2) Пропуски — порог «ID выше максимума» предполагает монотонный рост ID во времени внутри "
            "подкатегории; если это нарушится (новая карточка получит ID ниже уже зафиксированного максимума), "
            "метод её не заметит вовсе — это ограничение работает в обратную сторону от первого. "
            "Колонки «Протеин» и «Ингредиенты-клеймы» — только то, что вынесено крупно на лицевую сторону упаковки "
            "(не полный состав мелким шрифтом). «Завод-изготовитель» — юрлицо и адрес производства с этикетки "
            "(не бренд и не импортёр/представитель), где это удалось разобрать по фото.")
ws["A2"].font = note_font
ws.merge_cells("A2:N2")
ws["A2"].alignment = wrap
ws.row_dimensions[2].height = 108

headers = ["Дата обнаружения", "Метод обнаружения", "Вид животного", "Тип продукции", "Бренд", "Название товара",
           "Протеин", "Ингредиенты-клеймы (лицевая сторона упаковки)", "Завод-изготовитель", "ID на Ozon", "Ссылка", "Фото упаковки",
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
         plant="Не проверено отдельно; на лицевой стороне упаковки заявлено Made in Europe (страна не детализирована)"),
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
         plant="Таиланд (страна заявлена на лицевой стороне упаковки, конкретное юрлицо не проверялось)"),
    dict(animal="Кошки", ptype="Лакомства", brand="Titbit",
         name="Трава для кошек (для проращивания), 40 г х 12",
         oid="4098300679",
         url="https://www.ozon.ru/product/titbit-lakomstva-dlya-koshek-oves-dlya-prorashchivaniya-40-gr-12-up-4098300679/",
         photo="https://ir.ozone.ru/s3/multimedia-5/wc1000/6878482385.jpg",
         claims_short="100% натурально; немодифицированные семена; способствует пищеварению/выведению шерсти",
         protein="—(не мясной продукт, трава)", ingredients="Трава (овёс); витамины и фолиевая кислота (восполняет запас)",
         plant="Не найдено (только 2 фото в галерее, без этикетки)"),
    # --- Day 1 (2026-08-01): подтверждено сравнением ID с базой (реальные новые карточки, не сортировка) ---
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
         protein="Не указан на лицевой стороне упаковки", ingredients="—",
         plant="Farmina Pet Foods, Италия"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Farmina VetLife",
         name="Diabetic — диета для собак с диабетом, 6х300 г",
         oid="5327528788",
         url="https://www.ozon.ru/product/dieticheskiy-vlazhnyy-korm-dlya-sobak-farmina-vet-life-natural-diet-dog-diabetic-300h6-5327528788/",
         photo="https://ir.ozone.ru/s3/multimedia-1-l/wc1000/13388901681.jpg",
         claims_short="Поддерживает здоровое пищеварение; контроль уровня глюкозы в крови; подходит для собак с диабетом",
         protein="Не указан на лицевой стороне упаковки", ingredients="—",
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
         protein="Мясо (вид не конкретизирован на лицевой стороне упаковки)", ingredients="—",
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
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
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
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
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
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Farmina Pet Foods, Италия"),
    dict(animal="Кошки", ptype="Влажный корм", brand="AWARD",
         name="Sterilized, кусочки в соусе с индейкой, 85г х12",
         oid="5259703870",
         url="https://www.ozon.ru/product/5259703870/",
         photo="https://ir.ozone.ru/s3/multimedia-1-8/wc1000/8018668556.jpg",
         claims_short="100% BIO; Urinary Health; Nutraceutical Formula; Superpremium Cat Food",
         protein="Индейка", ingredients="—", plant="Не найдено"),

    # --- Day 5 (2026-08-04): подтверждено сравнением ID (макс. на 2026-08-03) ---
    # Дубли той же линейки, отличающиеся только фасовкой/весом или повторный листинг
    # того же товара другим продавцом — НЕ считаются отдельно (нет новой информации):
    # 5331256914 (тот же товар, что уже учтённый 5327553656 Farmina Gastrointestinal),
    # 5359501048, 5359513708 (дубли 5359453152 Puffins говядина+лосось по фасовке),
    # 5359688784, 5359678763 (дубли 5359660535 Puffins мясное ассорти по фасовке).
    dict(animal="Собаки", ptype="Сухой корм", brand="Pets Brunch",
         name="Adult Super Mini Breeds, индейка, гипоаллергенный, 4 кг",
         oid="5345830359",
         url="https://www.ozon.ru/product/suhoy-korm-pets-brunch-dlya-sobak-miniatyurnyh-porod-s-indeykoy-gipoallergennyy-4-kg-5345830359/",
         photo="https://ir.ozone.ru/s3/multimedia-1-4/wc1000/13460067904.jpg",
         claims_short="Sensitive; укрепляет иммунитет; защита пищеварения; красота кожи и шерсти; Gluten Free; 41% мясных ингредиентов; Belgian recipe",
         protein="Индейка", ingredients="Пребиотики",
         plant="Не найдено (Belgian recipe — рецептура, не завод)"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Mil.Dog",
         name="Полнорационный сухой корм для собак мелких пород, 10 кг",
         oid="5358300410",
         url="https://www.ozon.ru/product/korm-suhoy-10kg-dlya-sobak-melkih-porod-5358300410/",
         photo="https://ir.ozone.ru/s3/multimedia-1-o/wc1000/13502037768.jpg",
         claims_short="Сделано в России; без искусственных ароматизаторов, усилителей вкуса и сои",
         protein="Не конкретизирован (мясо и субпродукты)", ingredients="Таурин; Кальций/Фосфор; Витамин D; Жирные кислоты",
         plant="ООО «Ганеша» (Россия, Воронежская обл., г. Борисоглебск)"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Farmina N&D",
         name="Quail & Pumpkin, перепел и тыква, 285 г",
         oid="5357288481",
         url="https://www.ozon.ru/product/vlazhnyy-korm-farmina-n-d-quail-pumpkin-dlya-sobak-s-perepelom-i-tykvoy-285gr-5357288481/",
         photo="https://ir.ozone.ru/s3/multimedia-1-e/wc1000/13498610594.jpg",
         claims_short="Pumpkin line; Italian pet food",
         protein="Перепел", ingredients="Тыква",
         plant="Farmina Pet Foods, Италия"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Родные корма",
         name="Скоблянка мясная по-городецки, рубленое мясо, 970 г",
         oid="5358242931",
         url="https://www.ozon.ru/product/vlazhnyy-korm-rodnye-korma-konservy-dlya-sobak-skoblyanka-myasnaya-po-gorodetski-970-g-5358242931/",
         photo="https://ir.ozone.ru/s3/multimedia-1-0/wc1000/13451787720.jpg",
         claims_short="Премиум; мясо ингредиент №1; 75% мяса и мясных ингредиентов",
         protein="Говядина", ingredients="—", plant="Не найдено (Зоопторг — вероятно дистрибьютор, не завод)"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Родные корма",
         name="По-восточному, баранина с потрошками в желе, 970 г",
         oid="5358240926",
         url="https://www.ozon.ru/product/vlazhnyy-korm-rodnye-korma-konservy-dlya-sobak-po-vostochnomu-baranina-i-potroshki-zhele-970-g-5358240926/",
         photo="https://ir.ozone.ru/s3/multimedia-1-m/wc1000/13451786194.jpg",
         claims_short="Премиум; мясо ингредиент №1; 75% мяса и мясных ингредиентов",
         protein="Баранина", ingredients="—", plant="Не найдено (Зоопторг — вероятно дистрибьютор, не завод)"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Puffins",
         name="Индейка и лосось, для стерилизованных кошек, 400 г х2",
         oid="5359892602",
         url="https://www.ozon.ru/product/puffins-suhoy-korm-dlya-sterilizovannyh-koshek-i-kastrirovannyh-kotov-indeyka-i-losos-400-g-2-sht-5359892602/",
         photo="https://ir.ozone.ru/s3/multimedia-1-5/wc1000/13508291957.jpg",
         claims_short="Low Grain; здоровый иммунитет; контроль веса; здоровье мочевыдел. системы; здоровое пищеварение; здоровье шерсти и кожи",
         protein="Индейка, Лосось", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Puffins",
         name="Говядина и лосось, для взрослых кошек, 400 г",
         oid="5359453152",
         url="https://www.ozon.ru/product/puffins-suhoy-korm-dlya-koshek-govyadina-i-losos-400g-5359453152/",
         photo="https://ir.ozone.ru/s3/multimedia-1-n/wc1000/13506525671.jpg",
         claims_short="Low Grain; здоровый иммунитет; высокая усвояемость пищи; здоровое пищеварение; поддержание мышечного тонуса; здоровье кожи и шерсти",
         protein="Говядина, Лосось", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Puffins",
         name="Мясное ассорти, для взрослых кошек, 400 г",
         oid="5359660535",
         url="https://www.ozon.ru/product/puffins-suhoy-korm-dlya-vzroslyh-koshek-myasnoe-assorti-400-g-5359660535/",
         photo="https://ir.ozone.ru/s3/multimedia-1-m/wc1000/13507343950.jpg",
         claims_short="Low Grain; здоровый иммунитет; здоровье сердца и глаз; здоровое пищеварение; сильные мышцы и связки; здоровье шерсти и кожи",
         protein="Не конкретизирован (мясное ассорти)", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="GRANDORF",
         name="Lamb & Turkey Adult Indoor, ягнёнок с индейкой, Single Grain, 8 кг",
         oid="5359822345",
         url="https://www.ozon.ru/product/suhoy-korm-grandorf-yagnenok-s-indeykoy-dlya-domashnih-koshek-8-kg-5359822345/",
         photo="https://ir.ozone.ru/s3/multimedia-1-1/wc1000/13508049781.jpg",
         claims_short="Holistic & Hypoallergenic; Single Grain formula; Gluten Free; 70% premium animal ingredients; without plant protein; no artificial flavours/colours/preservatives; optimal Omega 3-6",
         protein="Ягнёнок, Индейка", ingredients="Омега-3/6",
         plant="Не найдено (не проверено отдельно для этого SKU)"),
    dict(animal="Кошки", ptype="Влажный корм", brand="DELISH",
         name="Super Premium, суп с курицей и овощами, 80 г",
         oid="5357563517",
         url="https://www.ozon.ru/product/konservy-dlya-koshek-delish-super-premium-sup-s-kuritsey-i-ovoshchami-80-gramm-5357563517/",
         photo="https://ir.ozone.ru/s3/multimedia-1-s/wc1000/13499540524.jpg",
         claims_short="Super Premium; высокое содержание чистого мяса; премиальный состав",
         protein="Курица", ingredients="Морковь; Овёс (пропаренный)", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="DELISH",
         name="Super Premium, суп с говядиной и овощами, 80 г",
         oid="5357853793",
         url="https://www.ozon.ru/product/konservy-dlya-koshek-delish-super-premium-sup-s-govyadinoy-i-ovoshchami-80-gramm-5357853793/",
         photo="https://ir.ozone.ru/s3/multimedia-1-s/wc1000/13500486352.jpg",
         claims_short="Super Premium; высокое содержание чистого мяса",
         protein="Говядина", ingredients="Морковь; Овёс (пропаренный)", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Blitz",
         name="Sterilised, индейка и клюква, паучи, 85 г х24",
         oid="5331257302",
         url="https://www.ozon.ru/product/pauchi-dlya-sterilizovannyh-koshek-indeyka-i-klyukva-85g-h-24-sht-5331257302/",
         photo="https://ir.ozone.ru/s3/multimedia-1-f/wc1000/13410175227.jpg",
         claims_short="Urinary care; Sensitive; обогащён пробиотиками; высокое качество",
         protein="Индейка", ingredients="Клюква; Пробиотики", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Mr.Whiskers",
         name="Крем-лакомство в стиках, ассорти 7 вкусов х2, 14х15 г",
         oid="5358325770",
         url="https://www.ozon.ru/product/krem-lakomstvo-dlya-koshek-v-stikah-5358325770/",
         photo="https://ir.ozone.ru/s3/multimedia-1-w/wc1000/13502125508.jpg",
         claims_short="Chicken+Catnip; Codfish+Fish oil (hair health); Salmon+Prebiotics (intestinal health); Rabbit (urinary tract); Tuna+Krill (immunity); Duck+Snow pear; Scallop",
         protein="Курица, Треска, Лосось, Кролик, Тунец, Утка, Гребешок",
         ingredients="Катнип (Chicken); Рыбий жир (Codfish); Пребиотики (Salmon); Криль (Tuna)",
         plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Мнямс",
         name="Ферма кота Федора, хрустящие подушечки с паштетом, тунец, 40 г",
         oid="5358242441",
         url="https://www.ozon.ru/product/lakomstvo-mnyams-hrustyashchie-podushechki-s-pashtetom-ferma-kota-fedora-dlya-koshek-i-kotyat-tunets-5358242441/",
         photo="https://ir.ozone.ru/s3/multimedia-1-p/wc1000/13492679965.jpg",
         claims_short="Без ГМО; без сои",
         protein="Тунец", ingredients="Таурин", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Мнямс",
         name="Ферма кота Федора, хрустящие подушечки с паштетом, говядина, 40 г",
         oid="5358240663",
         url="https://www.ozon.ru/product/lakomstvo-mnyams-hrustyashchie-podushechki-s-pashtetom-ferma-kota-fedora-dlya-koshek-i-5358240663/",
         photo="https://ir.ozone.ru/s3/multimedia-1-d/wc1000/13492670557.jpg",
         claims_short="Без ГМО; без сои",
         protein="Говядина", ingredients="Таурин", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Мнямс",
         name="Ферма кота Федора, хрустящие подушечки с паштетом, курица, 40 г",
         oid="5358241683",
         url="https://www.ozon.ru/product/lakomstvo-mnyams-hrustyashchie-podushechki-s-pashtetom-ferma-kota-fedora-dlya-koshek-i-5358241683/",
         photo="https://ir.ozone.ru/s3/multimedia-1-n/wc1000/13492672691.jpg",
         claims_short="Без ГМО; без сои",
         protein="Курица", ingredients="Таурин", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Четыре с хвостиком",
         name="Травка для кошек, 4 злаковые культуры, 80 г",
         oid="5358240549",
         url="https://www.ozon.ru/product/lakomstvo-chetyre-s-hvostikom-trava-dlya-prorashchivaniya-80-g-5358240549/",
         photo="https://ir.ozone.ru/s3/multimedia-1-l/wc1000/13450151757.jpg",
         claims_short="100% натуральный продукт; для выведения шерсти; источник витаминов А,В,Е",
         protein="—(не мясной продукт, трава)", ingredients="Овёс; Рожь; Пшеница; Тимофеевка; Витамины",
         plant="Не найдено (бренд АВЗ — юрлицо/адрес не подтверждены отдельно)"),

    # --- Day 6 (2026-08-12): подтверждено сравнением ID (макс. на 2026-08-04) ---
    # Дубли той же линейки, не считаются отдельно: 5423974817 (тот же товар, что 5430117217
    # АЛЬМАКОРМ говядина — идентичное фото листинга), 5429961720 (дубль 5429960930 Animonda
    # Carny по фасовке 400г вместо 200г). Также найден 5431960644 (трава для кошек, набор
    # для выращивания) — не профилируется отдельно: это не корм/лакомство, а сопутствующий
    # товар (семена в лотке), не входит в объём "клеймы на пачке корма/лакомства".
    dict(animal="Собаки", ptype="Каши", brand="АЛЬМАКОРМ",
         name="Каша для собак с говядиной, быстрого приготовления",
         oid="5430117217",
         url="https://www.ozon.ru/product/kasha-dlya-sobak-s-govyadinoy-bystrogo-prigotovleniya-5430117217/",
         photo="https://ir.ozone.ru/s3/multimedia-1-o/wc1000/8510341380.jpg",
         claims_short="100% натуральный, без вкусовых добавок; не требует варки, 5 минут и готово",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="АЛЬМАКОРМ",
         name="Каша для собак с индейкой, быстрого приготовления",
         oid="5423933879",
         url="https://www.ozon.ru/product/5423933879/",
         photo="https://ir.ozone.ru/s3/multimedia-1-w/wc1000/8566989188.jpg",
         claims_short="100% натуральный, без вкусовых добавок; не требует варки, 5 минут и готово",
         protein="Индейка", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="АЛЬМАКОРМ",
         name="Каша для собак свинина с говядиной, быстрого приготовления",
         oid="5423741999",
         url="https://www.ozon.ru/product/kasha-bystrogo-prigotovleniya-dlya-sobak-5kg-5423741999/",
         photo="https://ir.ozone.ru/s3/multimedia-1-i/wc1000/8567446878.jpg",
         claims_short="100% натуральный, без вкусовых добавок; не требует варки, 5 минут и готово",
         protein="Свинина, Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="СибКорм",
         name="Каша для собак с рыбой, 10 кг",
         oid="5432594014",
         url="https://www.ozon.ru/product/5432594014/",
         claims_short="Произведено на Алтае; без ГМО; из натуральных ингредиентов; для всех пород и возрастов",
         protein="Рыба", ingredients="—",
         plant="Алтайский край, Россия («Произведено на Алтае» — регион на упаковке, юрлицо не указано)"),
    dict(animal="Собаки", ptype="Каши", brand="Limuvil",
         name="Каша для собак, не требует варки",
         oid="5419467431",
         url="https://www.ozon.ru/product/5419467431/",
         claims_short="Не требует варки",
         protein="Не конкретизирован (заявлено «Злаки, Мясо», основной ингредиент — злаки)",
         ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="Командирская",
         name="Каша для собак «Командирская», 5 кг",
         oid="5418435027",
         url="https://www.ozon.ru/product/5418435027/",
         claims_short="С Алтайских полей; для всех пород и возрастов",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="Альма",
         name="Каша с мясом «Альма», с говядиной, быстрого приготовления",
         oid="5417460160",
         url="https://www.ozon.ru/product/kasha-s-govyadinoy-alma-korm-dlya-sobak-naturalnyy-5417460160/",
         claims_short="Без вкусовых добавок; не требует варки, 5 минут и готово; для крупных и средних пород",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Justin",
         name="Курица с печенью в соусе, кусочки, 75 г х 28 шт",
         oid="5432515655",
         url="https://www.ozon.ru/product/5432515655/",
         claims_short="Витамины и минералы; таурин; протеин",
         protein="Курица, Печень", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Мнямс",
         name="Микс кусочков в соусе для стерилизованных кошек, кролик и лосось, 85 г х 24 шт",
         oid="5431773123",
         url="https://www.ozon.ru/product/5431773123/",
         claims_short="Для стерилизованных кошек; профилактика мочекаменной болезни; контроль веса",
         protein="Кролик, Лосось", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Blitz",
         name="Sterilised, утка с лососем, нежные кусочки в соусе, 85 г х 24 шт",
         oid="5384510562",
         url="https://www.ozon.ru/product/5384510562/",
         claims_short="Holistic; 90% мясных ингредиентов в каждом кусочке",
         protein="Утка, Лосось", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Animonda Carny",
         name="Adult, курица, индейка и сердце утки, консервы, 200 г",
         oid="5429960930",
         url="https://www.ozon.ru/product/5429960930/",
         claims_short="100% свежие мясные ингредиенты; высокое содержание животного белка; без злаков; без добавления сахара; бренд с 1993 года",
         protein="Курица, Индейка, Утка", ingredients="—", plant="Германия"),
    dict(animal="Кошки", ptype="Влажный корм", brand="CLAN VET",
         name="Renal — профилактика болезней почек, пауч, 85 г",
         oid="5430728946",
         url="https://www.ozon.ru/product/5430728946/",
         claims_short="VET; профилактическая диета; профилактика болезней почек",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Royal Canin",
         name="Light Weight Care, в желе, паучи, 85 г х 16 шт",
         oid="5430727673",
         url="https://www.ozon.ru/product/5430727673/",
         claims_short="Помогает контролировать вес; свыше 90% кошек достигли более здорового веса за 8 недель; Health Nutrition с 1968 года",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Франция"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Animonda Vom Feinsten",
         name="Raffinesse, курица с нежной треской в желе, пауч, 24х85 г",
         oid="5430194418",
         url="https://www.ozon.ru/product/5430194418/",
         claims_short="Премиальная линейка «Raffinesse» (изысканность/утончённость)",
         protein="Курица, Треска", ingredients="—", plant="Германия"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Farmina N&D",
         name="Ocean, лосось, треска и дыня канталупа, беззерновой, 12 кг",
         oid="5424780704",
         url="https://www.ozon.ru/product/5424780704/",
         claims_short="Grain Free; Nutrition System for Carnivores; Italian formula",
         protein="Лосось, Треска", ingredients="Дыня канталупа", plant="Farmina Pet Foods, Италия"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Farmina N&D",
         name="Quinoa Weight Management, ягнёнок, киноа, брокколи и спаржа, 2,5 кг",
         oid="5424781094",
         url="https://www.ozon.ru/product/5424781094/",
         claims_short="Weight Management (контроль веса); Nutrition System for Carnivores; Italian formula",
         protein="Ягнёнок", ingredients="Киноа; Брокколи; Спаржа", plant="Farmina Pet Foods, Италия"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Farmina N&D",
         name="Ancestral Grain, курица, полба, овёс и гранат, для взрослых кошек, 10 кг",
         oid="5424780086",
         url="https://www.ozon.ru/product/5424780086/",
         claims_short="Ancestral Grain (древние злаки, низкозерновая рецептура); Nutrition System for Carnivores; Italian formula",
         protein="Курица", ingredients="Полба; Овёс; Гранат", plant="Farmina Pet Foods, Италия"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Farmina N&D",
         name="Ancestral Grain Neutered, курица, полба, овёс и гранат, для стерилизованных кошек, 10 кг",
         oid="5424781029",
         url="https://www.ozon.ru/product/5424781029/",
         claims_short="Ancestral Grain (древние злаки, низкозерновая рецептура); Nutrition System for Carnivores; Italian formula",
         protein="Курица", ingredients="Полба; Овёс; Гранат", plant="Farmina Pet Foods, Италия"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Farmina N&D",
         name="Ancestral Grain, ягнёнок, полба, овёс и черника, для взрослых кошек, 10 кг",
         oid="5424783286",
         url="https://www.ozon.ru/product/5424783286/",
         claims_short="Ancestral Grain (древние злаки, низкозерновая рецептура); Nutrition System for Carnivores; Italian formula",
         protein="Ягнёнок", ingredients="Полба; Овёс; Черника", plant="Farmina Pet Foods, Италия"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Farmina N&D",
         name="Pumpkin, перепел, тыква и гранат, беззерновой, для взрослых кошек, 5 кг",
         oid="5424776640",
         url="https://www.ozon.ru/product/5424776640/",
         claims_short="Grain Free; Nutrition System for Carnivores; Italian formula",
         protein="Перепел", ingredients="Тыква; Гранат", plant="Farmina Pet Foods, Италия"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Farmina N&D",
         name="Pumpkin, кабан, тыква и яблоко, беззерновой, 2,5 кг",
         oid="5424785573",
         url="https://www.ozon.ru/product/5424785573/",
         claims_short="Grain Free; Nutrition System for Carnivores; Italian formula",
         protein="Кабан", ingredients="Тыква; Яблоко", plant="Farmina Pet Foods, Италия"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Farmina N&D",
         name="Pumpkin Neutered, ягнёнок, тыква и черника, беззерновой, для стерилизованных кошек, 1,5 кг",
         oid="5424783287",
         url="https://www.ozon.ru/product/5424783287/",
         claims_short="Grain Free; Nutrition System for Carnivores; Italian formula",
         protein="Ягнёнок", ingredients="Тыква; Черника", plant="Farmina Pet Foods, Италия"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Farmina N&D",
         name="Quinoa Neutered Weight Management, утка, киноа, брокколи и спаржа, для стерилизованных собак, 12 кг",
         oid="5424784781",
         url="https://www.ozon.ru/product/5424784781/",
         claims_short="Weight Management (контроль веса); Nutrition System for Carnivores; Italian formula",
         protein="Утка", ingredients="Киноа; Брокколи; Спаржа", plant="Farmina Pet Foods, Италия"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Родные корма",
         name="Для взрослых стерилизованных кошек, курица и ягоды, 1,5 кг",
         oid="5658764220",
         url="https://www.ozon.ru/product/5658764220/",
         claims_short="Мясо ингредиент №1; 63% белка животного происхождения; Премиум; здоровая кожа и шерсть; контроль веса; комплекс пробиотиков",
         protein="Курица", ingredients="Клюква; Черника", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AWARD",
         name="Hairball & Indoor, утка с индейкой, зелёная чечевица и Юкка Шидигера, 10 кг",
         oid="5670474832",
         url="https://www.ozon.ru/product/5670474832/",
         claims_short="100% BIO; Hairball Control; Nutraceutical Formula; Superpremium Cat Food; Quality for Champions",
         protein="Утка, Индейка", ingredients="Зелёная чечевица; Юкка Шидигера", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AWARD",
         name="Urinary, курица, клюква, цикорий и рыбий жир, профилактика МКБ, 10 кг",
         oid="5670477242",
         url="https://www.ozon.ru/product/5670477242/",
         claims_short="100% BIO; Urinary Health (профилактика мочекаменной болезни); Nutraceutical Formula; Superpremium Cat Food; Quality for Champions",
         protein="Курица", ingredients="Клюква; Цикорий; Рыбий жир", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AWARD",
         name="Indoor Big cats, говядина с курицей, брокколи и юкка Шидигера, 10 кг",
         oid="5670498757",
         url="https://www.ozon.ru/product/5670498757/",
         claims_short="100% BIO; Nutraceutical Formula; Superpremium Cat Food; Quality for Champions",
         protein="Говядина, Курица", ingredients="Брокколи; Юкка Шидигера", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AWARD",
         name="Sterilized, индейка с курицей, клюква и цикорий, 10 кг",
         oid="5670499419",
         url="https://www.ozon.ru/product/5670499419/",
         claims_short="100% BIO; Urinary Health; Nutraceutical Formula; Superpremium Cat Food; Quality for Champions",
         protein="Индейка, Курица", ingredients="Клюква; Цикорий", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AWARD",
         name="Sterilized, белая рыба, семена льна, клюква и цикорий, 10 кг",
         oid="5670501608",
         url="https://www.ozon.ru/product/5670501608/",
         claims_short="100% BIO; Urinary Health; Nutraceutical Formula; Superpremium Cat Food; Quality for Champions",
         protein="Белая рыба", ingredients="Семена льна; Клюква; Цикорий", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Farmina N&D",
         name="Prime, курица и гранат, беззерновой, для котят, 5 кг",
         oid="5424782869",
         url="https://www.ozon.ru/product/5424782869/",
         claims_short="Grain Free; Nutrition System for Carnivores; Italian formula",
         protein="Курица", ingredients="Гранат", plant="Farmina Pet Foods, Италия"),
    dict(animal="Собаки", ptype="Лакомства", brand="TiTBiT",
         name="Fresh Snack, чек для мелких собак, мята/говядина, 5 шт х 55 г",
         oid="5428395063",
         url="https://www.ozon.ru/product/5428395063/",
         claims_short="Для освежения зубов; мята и хлорофилл; укрепление иммунитета; для дрессировки",
         protein="Говядина", ingredients="Мята; Хлорофилл; Ca/P", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="TiTBiT (Biff)",
         name="Dent, жевательный снек для чистки зубов, говядина, для мелких собак, 10 шт х 63 г",
         oid="5428925488",
         url="https://www.ozon.ru/product/5428925488/",
         claims_short="Профилактическое лакомство для чистки зубов; массирует дёсны; укрепление иммунитета; сделано в России",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="TiTBiT (Biff)",
         name="Dent, жевательный снек для чистки зубов, творог, для собак мини пород, 5 шт х 30 г",
         oid="5429259395",
         url="https://www.ozon.ru/product/5429259395/",
         claims_short="Профилактическое лакомство для чистки зубов; массирует дёсны; укрепление иммунитета",
         protein="—(немясной протеин: творог)", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="TiTBiT (Biff)",
         name="Dent, жевательный снек для чистки зубов, телятина, для собак мини пород, 7 шт х 20 г",
         oid="5431102190",
         url="https://www.ozon.ru/product/5431102190/",
         claims_short="Профилактическое лакомство для чистки зубов; массирует дёсны; укрепление иммунитета",
         protein="Телятина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Yummy",
         name="Мясной Деликатес, индейка в желе, консервы, 100 г",
         oid="5669273502",
         url="https://www.ozon.ru/product/5669273502/",
         claims_short="Без зерна; монобелковый",
         protein="Индейка", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Мнямс",
         name="Максимум вкуса, мясное рагу, говядина с морковью, консервы, 85 г х 12",
         oid="5670094001",
         url="https://www.ozon.ru/product/5670094001/",
         claims_short="Не содержит злаков",
         protein="Говядина", ingredients="Морковь", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Мнямс",
         name="Максимум вкуса, томлёная телятина в соусе, пауч, 85 г",
         oid="5670712472",
         url="https://www.ozon.ru/product/5670712472/",
         claims_short="Не содержит злаков",
         protein="Телятина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Мнямс",
         name="Сильный иммунитет, паштет из утки и перепела, для всех пород, 200 г х 3",
         oid="5670714813",
         url="https://www.ozon.ru/product/5670714813/",
         claims_short="Не содержит злаков; сильный иммунитет",
         protein="Утка, Перепел", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Мнямс",
         name="Мусс из курицы с печенью, дополнительный корм, 80 г х 6",
         oid="5670718836",
         url="https://www.ozon.ru/product/5670718836/",
         claims_short="Не содержит злаков",
         protein="Курица, Печень", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Мнямс",
         name="Террин по-версальски, паштет из телятины с ветчиной, 200 г х 5",
         oid="5670720625",
         url="https://www.ozon.ru/product/5670720625/",
         claims_short="Не содержит злаков",
         protein="Телятина", ingredients="Ветчина", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Мнямс",
         name="Мусс из утки, дополнительный корм, 80 г х 2",
         oid="5671464738",
         url="https://www.ozon.ru/product/5671464738/",
         claims_short="Не содержит злаков",
         protein="Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Мнямс",
         name="Мусс по-французски, паштет из индейки с цукини, для щенков до 4 месяцев, 200 г х 3",
         oid="5671505639",
         url="https://www.ozon.ru/product/5671505639/",
         claims_short="Не содержит злаков",
         protein="Индейка", ingredients="Цукини", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="AWARD",
         name="Monoprotein, паштет, телятина, консервы, 400 г х 12",
         oid="5430211436",
         url="https://www.ozon.ru/product/5430211436/",
         claims_short="Monoprotein; 100% BIO; Superpremium Dog Food; Quality for Champions",
         protein="Телятина (по названию/этикетке; на фото товара — тот же шаблонный паштет «Lamb & Apple», что и у др. вкусов линейки)",
         ingredients="—",
         plant="Не проверено отдельно; вероятно тот же завод, что у AWARD сухой корм (ООО «ПетЭксперт», Липецкая обл.)"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Karmy",
         name="Starter, для щенков и кормящих/берем. сук всех пород, индейка",
         oid="5432604231",
         url="https://www.ozon.ru/product/5432604231/",
         claims_short="Способствует формированию пищеварительной системы; поддерживает работу пищевыделительной системы; поддерживает здоровье иммунной системы; свежее мясо",
         protein="Индейка", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Farmina N&D",
         name="Prime, ягнёнок и черника, беззерновой, для взрослых собак, 12 кг",
         oid="5623144132",
         url="https://www.ozon.ru/product/5623144132/",
         claims_short="Grain Free; Nutrition System for Carnivores; Italian formula",
         protein="Ягнёнок", ingredients="Черника", plant="Farmina Pet Foods, Италия"),
    dict(animal="Собаки", ptype="Сухой корм", brand="GRANDORF",
         name="Veterinary Diet Hypoallergenic, кролик, при пищевой аллергии, 3 кг",
         oid="5634101309",
         url="https://www.ozon.ru/product/5634101309/",
         claims_short="Ветеринарная диета (Approved by Vets); Hypoallergenic (пищевая аллергия и непереносимость); Gluten & Grain Free; антистресс; пробиотики/постбиотики для ЖКТ и иммунитета; здоровье кожи и шерсти",
         protein="Кролик", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="GRANDORF",
         name="Fresh, филе лосося с бататом, беззерновой холистик, для мелких пород, 1 кг",
         oid="5662623115",
         url="https://www.ozon.ru/product/5662623115/",
         claims_short="Holistic; Grain Free; живые пробиотики",
         protein="Лосось", ingredients="Батат", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="GRANDORF",
         name="Fresh, мясо индейки с бататом, беззерновой холистик, для средних и крупных пород, 10 кг",
         oid="5662623376",
         url="https://www.ozon.ru/product/5662623376/",
         claims_short="Holistic; Grain Free; живые пробиотики",
         protein="Индейка", ingredients="Батат", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="GRANDORF",
         name="Fresh, мясо утки с бататом, беззерновой холистик, для средних и крупных пород, 10 кг",
         oid="5662626493",
         url="https://www.ozon.ru/product/5662626493/",
         claims_short="Holistic; Grain Free; живые пробиотики",
         protein="Утка", ingredients="Батат", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Royal Canin",
         name="Expert Neutered Adult, для стерилизованных/кастрированных собак мелких пород, 3,5 кг",
         oid="5670255238",
         url="https://www.ozon.ru/product/5670255238/",
         claims_short="№1 по рекомендации ветеринарных врачей; ветеринарный рацион (Expert)",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="AWARD",
         name="HYPO, белая рыба, брокколи, сельдерей и семена льна, для взрослых собак всех пород, 12 кг",
         oid="5670469697",
         url="https://www.ozon.ru/product/5670469697/",
         claims_short="Hypoallergenic; 100% BIO; 65% животного белка; 0% глютен/соя; Long-term Energy Boost; Nutraceutical Formula; Superpremium Dog Food; Quality for Champions",
         protein="Белая рыба", ingredients="Брокколи; Сельдерей; Семена льна", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Pedigree",
         name="Dentastix Light, для мелких пород, индейка, 35 палочек",
         oid="5432823007",
         url="https://www.ozon.ru/product/5432823007/",
         claims_short="Снижает образование зубного камня; ежедневный уход за зубами; разработано с ветеринарами; на 30% меньше калорий",
         protein="Индейка", ingredients="—", plant="Не найдено"),

    # --- Доразбор 12.08: найдены аудитом ?sorting=new (ID выше порога Day 6, пропущены
    # при первом проходе по подкатегории). Дубль 5424780528 (Farmina N&D Quinoa Neutered,
    # утка, 12 кг) — идентичное фото и рецептура, что у уже учтённого 5424784781 —
    # не профилируется отдельно, только в истории ID.
    dict(animal="Собаки", ptype="Влажный корм", brand="CLAN VET",
         name="Gastrointestinal — профилактика болезней ЖКТ, пауч, 85 г х 14",
         oid="5433103013",
         url="https://www.ozon.ru/product/5433103013/",
         claims_short="VET; профилактическая диета; профилактика болезней ЖКТ",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="CLAN Classic",
         name="Кусочки в соусе, кролик, консервы",
         oid="5433102619",
         url="https://www.ozon.ru/product/5433102619/",
         claims_short="Натуральные ингредиенты в составе; клюква — профилактика МКБ; ромашка — улучшает пищеварение",
         protein="Кролик", ingredients="Клюква; Брокколи; Ромашка", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Berkley",
         name="№4, перепёлка с курицей и овощами, консервы, 200 г х 6",
         oid="5433103250",
         url="https://www.ozon.ru/product/5433103250/",
         claims_short="100% Natural Product; Local",
         protein="Перепел, Курица", ingredients="Овощи", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Berkley",
         name="№6, кролик с говядиной и болгарским перцем, консервы, 200 г х 6",
         oid="5433103558",
         url="https://www.ozon.ru/product/5433103558/",
         claims_short="100% Natural Product; Local",
         protein="Кролик, Говядина", ingredients="Болгарский перец", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Berkley",
         name="№7, индейка с яблоками и отрубями, консервы, 200 г х 6",
         oid="5433104303",
         url="https://www.ozon.ru/product/5433104303/",
         claims_short="100% Natural Product; Local",
         protein="Индейка", ingredients="Яблоки; Отруби", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Delish",
         name="Консервы для собак, кролик, 400 г х 6",
         oid="5433102991",
         url="https://www.ozon.ru/product/5433102991/",
         claims_short="Натуральный состав; 100% мяса; без ГМО",
         protein="Кролик", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="CLAN PRIDE",
         name="Калтык и язык, консервы, 340 г х 6",
         oid="5433102865",
         url="https://www.ozon.ru/product/5433102865/",
         claims_short="Без заявленных клеймов — только состав и пищевая ценность на этикетке",
         protein="Калтык и язык (субпродукты)", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Delish",
         name="Консервы для собак, говядина и ягнёнок, 400 г х 6",
         oid="5433103171",
         url="https://www.ozon.ru/product/5433103171/",
         claims_short="Натуральный состав; 100% мяса; без ГМО",
         protein="Говядина, Ягнёнок", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="INABA",
         name="Neji, говядина и куриное филе на гриле со вкусом сыра, 8 г х 3",
         oid="5632614860",
         url="https://www.ozon.ru/product/5632614860/",
         claims_short="Вкусовая привлекательность (гриль, вкус сыра)",
         protein="Говядина, Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="INABA",
         name="Churu Bee, запечённые трубочки, куриное филе, для поддержания здоровья суставов, 10 г х 3",
         oid="5632615231",
         url="https://www.ozon.ru/product/5632615231/",
         claims_short="Поддержание здоровья суставов",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="INABA",
         name="Churu, лакомство-пюре, куриное филе, для здоровья кожи и шерсти, 14 г х 4",
         oid="5632624539",
         url="https://www.ozon.ru/product/5632624539/",
         claims_short="Здоровье кожи и шерсти",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="INABA",
         name="Churu Ball, шарики с кремовой начинкой, куриное филе с добавлением овощей, 10 г х 3",
         oid="5654589093",
         url="https://www.ozon.ru/product/5654589093/",
         claims_short="Вкусовая привлекательность (кремовая начинка)",
         protein="Курица", ingredients="Овощи", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="INABA",
         name="Churu, лакомство-пюре, куриное филе, для здоровья ЖКТ, 14 г х 4",
         oid="5654591896",
         url="https://www.ozon.ru/product/5654591896/",
         claims_short="Здоровье ЖКТ",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="INABA",
         name="Churu Ball, шарики с кремовой начинкой, куриное филе, 10 г х 3",
         oid="5654592771",
         url="https://www.ozon.ru/product/5654592771/",
         claims_short="Вкусовая привлекательность (кремовая начинка)",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="INABA",
         name="Tobitsuku, запечённое лакомство, куриное филе, для поддержания здоровья кожи, 25 г",
         oid="5654599026",
         url="https://www.ozon.ru/product/5654599026/",
         claims_short="Поддержание здоровья кожи",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="INABA",
         name="Tobitsuku, запечённое лакомство, куриное филе, для поддержания здоровья ЖКТ, 25 г",
         oid="5654599608",
         url="https://www.ozon.ru/product/5654599608/",
         claims_short="Поддержание здоровья ЖКТ",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Мнямс",
         name="DENTAL, зубной снек с лососем, 85 г",
         oid="5670711483",
         url="https://www.ozon.ru/product/5670711483/",
         claims_short="Профилактика образования зубного камня; массаж дёсен",
         protein="Лосось", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Мнямс",
         name="Колбаски аппетитные, 80 г",
         oid="5670715414",
         url="https://www.ozon.ru/product/5670715414/",
         claims_short="100% натуральное лакомство; без добавок и красителей; улучшает пищеварение",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Мнямс",
         name="Грудки куриные, 50 г",
         oid="5670717416",
         url="https://www.ozon.ru/product/5670717416/",
         claims_short="100% натуральное лакомство; без добавок и красителей; улучшает пищеварение",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Мнямс",
         name="Филе миньон по-фламандски, 75 г х 2",
         oid="5670718664",
         url="https://www.ozon.ru/product/5670718664/",
         claims_short="100% натуральное лакомство; без добавок и красителей; улучшает пищеварение",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Мнямс",
         name="Рубец говяжий, 70 г х 2",
         oid="5670720986",
         url="https://www.ozon.ru/product/5670720986/",
         claims_short="100% натуральное лакомство; без добавок и красителей; улучшает пищеварение",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Мнямс",
         name="DENTAL, зубной снек с говядиной, для крупных пород",
         oid="5671459555",
         url="https://www.ozon.ru/product/5671459555/",
         claims_short="Профилактика образования зубного камня; массаж дёсен",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Мнямс",
         name="DENTAL, зубные косточки, для щенков и взрослых животных, 70 г",
         oid="5671490575",
         url="https://www.ozon.ru/product/5671490575/",
         claims_short="Профилактика образования зубного камня; массаж дёсен",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Мнямс",
         name="Трахея говяжья, 95 г",
         oid="5671509904",
         url="https://www.ozon.ru/product/5671509904/",
         claims_short="100% натуральное лакомство; без добавок и красителей; улучшает пищеварение",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Мнямс",
         name="Колбаски из баранины, 125 г х 2",
         oid="5671561038",
         url="https://www.ozon.ru/product/5671561038/",
         claims_short="100% натуральное лакомство; без добавок и красителей; улучшает пищеварение",
         protein="Баранина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Кинолог в берцах",
         name="Утиные головы для собак, натуральное лакомство, 50 г",
         oid="5436654518",
         url="https://www.ozon.ru/product/5436654518/",
         claims_short="100% утка, монопротеин; чистые зубы и здоровые дёсны; снимает стресс; источник белка и аминокислот; без красителей, консервантов и усилителей вкуса",
         protein="Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Farmina N&D",
         name="Prime, курица и гранат, беззерновой, для взрослых собак, 2,5 кг",
         oid="5424786354",
         url="https://www.ozon.ru/product/5424786354/",
         claims_short="Grain Free; Nutrition System for Carnivores; Italian formula",
         protein="Курица", ingredients="Гранат", plant="Farmina Pet Foods, Италия"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Farmina N&D",
         name="Ancestral Grain, ягнёнок, полба, овёс и черника, для взрослых собак, 12 кг",
         oid="5424780417",
         url="https://www.ozon.ru/product/5424780417/",
         claims_short="Ancestral Grain (древние злаки, низкозерновая рецептура); Nutrition System for Carnivores; Italian formula",
         protein="Ягнёнок", ingredients="Полба; Овёс; Черника", plant="Farmina Pet Foods, Италия"),

    # =====================================================================
    # Day 7 (2026-08-24): подтверждено сравнением ID (макс. на 2026-08-12).
    # Разбор велся ускоренно (шаблонизация линеек по просьбе пользователя) —
    # для линеек с несколькими вкусами полностью проверено фото одного
    # представителя, остальные добавлены по названию карточки; дубли по
    # фасовке/повторному листингу не профилируются отдельно.
    # =====================================================================

    # --- Каша для собак ---
    dict(animal="Собаки", ptype="Каши", brand="АЛЬМАКОРМ",
         name="Каша для собак с курицей, быстрого приготовления",
         oid="5542503462",
         url="https://www.ozon.ru/product/5542503462/",
         claims_short="100% натуральный, без вкусовых добавок; не требует варки, 5 минут и готово",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="АЛЬМАКОРМ",
         name="Каша для собак с рыбой, быстрого приготовления",
         oid="5445105149",
         url="https://www.ozon.ru/product/5445105149/",
         claims_short="100% натуральный, без вкусовых добавок; не требует варки, 5 минут и готово",
         protein="Рыба", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="АЛЬМАКОРМ",
         name="Каша для собак с гречкой, быстрого приготовления",
         oid="5445127064",
         url="https://www.ozon.ru/product/5445127064/",
         claims_short="100% натуральный, без вкусовых добавок; не требует варки, 5 минут и готово",
         protein="Не мясной продукт (гречневая крупа)", ingredients="Гречка", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="ЭТИ ЗВЕРИ",
         name='Каша быстрого приготовления "Лапа Злак", 400 г',
         oid="5456129061",
         url="https://www.ozon.ru/product/5456129061/",
         claims_short="Натуральный продукт; не требует варки; сделано в России",
         protein="Не мясной продукт (злаковая смесь)", ingredients="Злаки", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="ЭТИ ЗВЕРИ",
         name="Каша быстрого приготовления, рисовая, 400 г",
         oid="5456129661",
         url="https://www.ozon.ru/product/5456129661/",
         claims_short="Натуральный продукт; не требует варки; сделано в России",
         protein="Не мясной продукт (рис)", ingredients="Рис", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="ЭТИ ЗВЕРИ",
         name="Каша быстрого приготовления, злаковая, 400 г",
         oid="5456131380",
         url="https://www.ozon.ru/product/5456131380/",
         claims_short="Натуральный продукт; не требует варки; сделано в России",
         protein="Не мясной продукт (злаковая смесь)", ingredients="Злаки", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="ЭТИ ЗВЕРИ",
         name="Каша быстрого приготовления, гречневая, 400 г",
         oid="5456131513",
         url="https://www.ozon.ru/product/5456131513/",
         claims_short="Натуральный продукт; не требует варки; сделано в России",
         protein="Не мясной продукт (гречневая крупа)", ingredients="Гречка", plant="Не найдено"),
    dict(animal="Собаки", ptype="Каши", brand="СибКорм",
         name="Каша для собак с мясом, 10 кг",
         oid="5479813617",
         url="https://www.ozon.ru/product/5479813617/",
         claims_short="Произведено на Алтае; 100% натуральный состав; без ГМО",
         protein="Мясо (вид не конкретизирован на лицевой стороне упаковки)", ingredients="—",
         plant="Алтайский край, Россия («Произведено на Алтае» — регион на упаковке, юрлицо не указано)"),
    dict(animal="Собаки", ptype="Каши", brand="СибКорм",
         name="Каша для собак Ассорти (мясо + рыба), 2х10 кг",
         oid="5500647058",
         url="https://www.ozon.ru/product/5500647058/",
         claims_short="Произведено на Алтае; 100% натуральный состав; без ГМО",
         protein="Мясо + Рыба (ассорти, без детализации на упаковке)", ingredients="—",
         plant="Алтайский край, Россия («Произведено на Алтае» — регион на упаковке, юрлицо не указано)"),

    # --- Day 7 (2026-08-24): Сухой корм для собак ---
    dict(animal="Собаки", ptype="Сухой корм", brand="Wanpy Dog",
         name="Беззерновой, курица с черникой и горошком (Oven Baked Bites), 500 г",
         oid="5515836917",
         url="https://www.ozon.ru/product/5515836917/",
         claims_short="Grain-Free; 84% Animal Protein; Healthy Digestion; No added corn/wheat/soy/lentils",
         protein="Курица", ingredients="Черника; Горошек", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Wanpy Dog",
         name="Беззерновой, с лососем, 1,5 кг",
         oid="5515838796",
         url="https://www.ozon.ru/product/5515838796/",
         claims_short="Grain-Free; 84% Animal Protein; Healthy Digestion; No added corn/wheat/soy/lentils",
         protein="Лосось", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Wanpy Dog",
         name="Беззерновой, с ягнёнком, 500 г",
         oid="5524575556",
         url="https://www.ozon.ru/product/5524575556/",
         claims_short="Grain-Free; 84% Animal Protein; Healthy Digestion; No added corn/wheat/soy/lentils",
         protein="Ягнёнок", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="SIRIUS",
         name="Premium, индейка с овощами, для взрослых собак крупных пород, 15 кг",
         oid="5529408722",
         url="https://www.ozon.ru/product/5529408722/",
         claims_short="41% животных ингредиентов; комплекс пробиотиков ProStor; крепкие кости; поддерживает иммунитет; здоровое сердце; Premium Food",
         protein="Индейка", ingredients="Морковь; Сельдерей; Пробиотики", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Provipet",
         name="С ягнёнком, для собак всех пород, 10 кг",
         oid="5530070606",
         url="https://www.ozon.ru/product/5530070606/",
         claims_short="100% полноценное сбалансированное питание; высокая усвояемость; отличное пищеварение; крепкие зубы и кости; устойчивая иммунная система; здоровье кожи и шерсти; без искусственных ароматизаторов и усилителей вкуса; сделано в России",
         protein="Ягнёнок", ingredients="—", plant="Не найдено (заявлено «Сделано в России», юрлицо не указано)"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Provipet",
         name="С индейкой, для собак всех пород",
         oid="5555438805",
         url="https://www.ozon.ru/product/5555438805/",
         claims_short="100% полноценное сбалансированное питание; высокая усвояемость; отличное пищеварение; крепкие зубы и кости; устойчивая иммунная система; здоровье кожи и шерсти; без искусственных ароматизаторов и усилителей вкуса; сделано в России",
         protein="Индейка", ingredients="—", plant="Не найдено (заявлено «Сделано в России», юрлицо не указано)"),
    dict(animal="Собаки", ptype="Сухой корм", brand="AWARD",
         name="Monoprotein, лосось с семенами льна и алоэ вера, для мелких пород, 1,5 кг",
         oid="5542850495",
         url="https://www.ozon.ru/product/5542850495/",
         claims_short="Hypo & Digestion; Low Grain; 100% BIO; Fresh Meat; Superpremium Dog Food",
         protein="Лосось", ingredients="Семена льна; Алоэ вера", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="AlphaPet",
         name="Superpremium Monoprotein, из белой рыбы, для взрослых собак средних и крупных пород, 12 кг",
         oid="5556742789",
         url="https://www.ozon.ru/product/5556742789/",
         claims_short="Монобелковая рецептура; 30% свежей белой рыбы (ингредиент №1); чувствительное пищеварение; здоровая кожа и блестящая шерсть; Superpremium",
         protein="Белая рыба", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Grandorf",
         name="Fresh, индейка с бататом, для взрослых собак средних и крупных пород, 1 кг",
         oid="5566163855",
         url="https://www.ozon.ru/product/5566163855/",
         claims_short="Holistic&Hypoallergenic; Gut Health/пробиотики; Omega-3/6; Grain Free; мясо первым ингредиентом",
         protein="Индейка", ingredients="Батат", plant="Не проверено отдельно; вероятно United Petfood (та же группа, что и другие GRANDORF Fresh)"),
    dict(animal="Собаки", ptype="Сухой корм", brand="AlphaPet Menu",
         name="Adult All Breeds, с говядиной, свежее мясо, 2 кг",
         oid="5566490497",
         url="https://www.ozon.ru/product/5566490497/",
         claims_short="Со свежим мясом; 100% натуральные ингредиенты; сбалансированный рацион",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Vivere",
         name="Maxi Adult, ягнёнок, низкозерновой, для крупных пород, 12 кг",
         oid="5568844559",
         url="https://www.ozon.ru/product/5568844559/",
         claims_short="Holistic Philosophy; 100% Natural Pet Food; Low Grain; 60% Animal Ingredients",
         protein="Ягнёнок", ingredients="Рис; Горох", plant="Италия"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Purina Pro Plan",
         name="Veterinary Diets EN Gastrointestinal, для щенков и взрослых собак при расстройствах пищеварения, 1,5 кг х 2",
         oid="5574603661",
         url="https://www.ozon.ru/product/5574603661/",
         claims_short="VET; ветеринарная диета при расстройствах пищеварения",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Сухой корм", brand="Hilton",
         name="Гипоаллергенный, говядина, для крупных пород, 2х10 кг",
         oid="5578347424",
         url="https://www.ozon.ru/product/5578347424/",
         claims_short="Гипоаллергенный корм",
         protein="Говядина", ingredients="—", plant="Не найдено"),

    # --- Day 7 (2026-08-24): Влажный корм для собак ---
    dict(animal="Собаки", ptype="Влажный корм", brand="Мнямс",
         name="Максимум вкуса, томленая телятина в соусе, пауч 85 г",
         oid="5471075781",
         url="https://www.ozon.ru/product/5471075781/",
         claims_short="Без консервантов; без ГМО",
         protein="Телятина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Wanpy Dog",
         name="Консервы для собак, курица, 375 г х 6",
         oid="5530369991",
         url="https://www.ozon.ru/product/5530369991/",
         claims_short="Premium Quality; Abundant Nutrition; Strengthens Bones; Healthy Digestion",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Wanpy Dog",
         name="Консервы для собак, говядина, 375 г х 6",
         oid="5530650366",
         url="https://www.ozon.ru/product/5530650366/",
         claims_short="Premium Quality; Abundant Nutrition; Strengthens Bones; Healthy Digestion",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Wanpy Dog",
         name="Консервы для собак, ягнёнок, 375 г х 6",
         oid="5530650986",
         url="https://www.ozon.ru/product/5530650986/",
         claims_short="Premium Quality; Abundant Nutrition; Strengthens Bones; Healthy Digestion",
         protein="Ягнёнок", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Royal Canin",
         name="Mini Adult, паштет, для взрослых собак мелких пород, 195 г х 6",
         oid="5544614194",
         url="https://www.ozon.ru/product/5544614194/",
         claims_short="Сделано в Австрии",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Австрия"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Pedigree",
         name="Для щенков от 1 месяца, паштет с телятиной, 28 шт по 80 г",
         oid="5549300374",
         url="https://www.ozon.ru/product/5549300374/",
         claims_short="Натуральное мясо в составе",
         protein="Телятина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Зоогурман",
         name="Большая миска, консервы, индейка, 970 г",
         oid="5556733970",
         url="https://www.ozon.ru/product/5556733970/",
         claims_short="Без заявленных клеймов — только состав и вес на этикетке",
         protein="Индейка", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Влажный корм", brand="Зоогурман",
         name="Большая миска, консервы, телятина с рубцом, 970 г",
         oid="5556739255",
         url="https://www.ozon.ru/product/5556739255/",
         claims_short="Без заявленных клеймов — только состав и вес на этикетке",
         protein="Телятина, Рубец", ingredients="—", plant="Не найдено"),

    # --- Day 7 (2026-08-24): Лакомства для собак ---
    dict(animal="Собаки", ptype="Лакомства", brand="—",
         name="Лёгкое говяжье для дрессировки собак, 1000 г",
         oid="5544395456",
         url="https://www.ozon.ru/product/5544395456/",
         claims_short="100% натуральное; поддержка зубов и дёсен",
         protein="Говядина (лёгкое)", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="biohaus",
         name="Медальоны из утки, 750 г",
         oid="5555142584",
         url="https://www.ozon.ru/product/5555142584/",
         claims_short="100% натуральный продукт; без ГМО; без глютена; 99% владельцев рекомендуют",
         protein="Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Lucky Animals",
         name="Luckers, яблоко с говядиной, 8 шт",
         oid="5555380815",
         url="https://www.ozon.ru/product/5555380815/",
         claims_short="100% натуральное лакомство; 100% говядина; сделано вручную",
         protein="Говядина", ingredients="Яблоко", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="biohaus",
         name="Суши из кролика и трески, 750 г",
         oid="5555663607",
         url="https://www.ozon.ru/product/5555663607/",
         claims_short="100% натуральный продукт; без ГМО; без глютена; 99% владельцев рекомендуют",
         protein="Кролик, Треска", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="biohaus",
         name="Суши из индейки и трески, 150 г",
         oid="5556221688",
         url="https://www.ozon.ru/product/5556221688/",
         claims_short="100% натуральный продукт; без ГМО; без глютена; 99% владельцев рекомендуют",
         protein="Индейка, Треска", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="biohaus",
         name="Шашлычки куриные на жевательной шпажке, 2х95 г",
         oid="5556538078",
         url="https://www.ozon.ru/product/5556538078/",
         claims_short="100% натуральный продукт; без ГМО; без глютена; 99% владельцев рекомендуют",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="biohaus",
         name="Уши кролика с мясом утки, 2х75 г",
         oid="5556653006",
         url="https://www.ozon.ru/product/5556653006/",
         claims_short="100% натуральный продукт; без ГМО; без глютена; 99% владельцев рекомендуют",
         protein="Кролик, Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="TiTBiT",
         name="Мясные сушки с говядиной, 1 кг",
         oid="5577356638",
         url="https://www.ozon.ru/product/5577356638/",
         claims_short="Мясное печенье; 50% мясопродуктов по сырому весу; сделано в России",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="TiTBiT",
         name="Мясное печенье «Бискотти», 1 кг",
         oid="5577356979",
         url="https://www.ozon.ru/product/5577356979/",
         claims_short="Мясное печенье; 50% мясопродуктов по сырому весу; сделано в России",
         protein="Не конкретизирован на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="TiTBiT (Biff)",
         name="Dent, жевательный снек для чистки зубов, сыр, 20 шт",
         oid="5577357024",
         url="https://www.ozon.ru/product/5577357024/",
         claims_short="Профилактическое лакомство для чистки зубов; массирует дёсны; укрепление иммунитета; для дрессировки",
         protein="—(немясной протеин: сыр)", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="TiTBiT (Biff)",
         name="Dent, жевательный снек для чистки зубов, утка, 14 шт",
         oid="5577357304",
         url="https://www.ozon.ru/product/5577357304/",
         claims_short="Профилактическое лакомство для чистки зубов; массирует дёсны; укрепление иммунитета; для дрессировки",
         protein="Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="TiTBiT",
         name="Соломка из лосося, рыбные лакомства, 80 г",
         oid="5577357359",
         url="https://www.ozon.ru/product/5577357359/",
         claims_short="100% натурально; содержит Омега-3 и 6; полезно для шерсти и кожи; сделано в России",
         protein="Лосось", ingredients="Омега-3/6", plant="Не найдено (заявлено «Сделано в России»)"),
    dict(animal="Собаки", ptype="Лакомства", brand="TiTBiT",
         name="Кубики из лосося, рыбные лакомства, 80 г",
         oid="5577358728",
         url="https://www.ozon.ru/product/5577358728/",
         claims_short="100% натурально; содержит Омега-3 и 6; полезно для шерсти и кожи; сделано в России",
         protein="Лосось", ingredients="Омега-3/6", plant="Не найдено (заявлено «Сделано в России»)"),
    dict(animal="Собаки", ptype="Лакомства", brand="TiTBiT",
         name="Крекер говяжий XL",
         oid="5577359119",
         url="https://www.ozon.ru/product/5577359119/",
         claims_short="Способствует удалению зубного налёта и укреплению дёсен; 100% натуральный продукт, без искусственных красителей, консервантов и ароматизаторов",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="Мистер Грызкин",
         name="Говяжья трахея, 5-7 см, 200 г",
         oid="5577402177",
         url="https://www.ozon.ru/product/5577402177/",
         claims_short="Полностью натурально",
         protein="Говядина (трахея)", ingredients="—", plant="Не найдено"),
    dict(animal="Собаки", ptype="Лакомства", brand="biohaus",
         name="Сэндвич с завязанной уткой, 2х75 г",
         oid="5585499483",
         url="https://www.ozon.ru/product/5585499483/",
         claims_short="100% натуральный продукт; без ГМО; без глютена; 99% владельцев рекомендуют",
         protein="Утка", ingredients="—", plant="Не найдено"),

    # --- Day 7 (2026-08-24): Сухой корм для кошек ---
    dict(animal="Кошки", ptype="Сухой корм", brand="Wanpy Cat",
         name="Беззерновой, с курицей, для котят, 1,5 кг",
         oid="5515840508",
         url="https://www.ozon.ru/product/5515840508/",
         claims_short="Grain-Free; 84% Animal Protein; Healthy Digestion; No added corn/wheat/soy/lentils",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Wanpy Cat",
         name="Беззерновой, с лососем, для взрослых кошек, 1,5 кг",
         oid="5515844102",
         url="https://www.ozon.ru/product/5515844102/",
         claims_short="Grain-Free; 84% Animal Protein; Healthy Digestion; No added corn/wheat/soy/lentils",
         protein="Лосось", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AWARD",
         name="Monoprotein, для стерилизованных кошек, лосось со спаржей и ацеролой, 1,5 кг",
         oid="5542531112",
         url="https://www.ozon.ru/product/5542531112/",
         claims_short="Hypo & Digestion; Low Grain; 100% BIO; Fresh Meat; Superpremium Cat Food",
         protein="Лосось", ingredients="Спаржа; Ацерола", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AJO",
         name="Special Line Delicate Taste, для привередливых кошек и котят, с индейкой, 0,4 кг",
         oid="5542594928",
         url="https://www.ozon.ru/product/5542594928/",
         claims_short="Holistic View; Low Grain; Wheat/Corn Free",
         protein="Индейка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AlphaPet",
         name="Для взрослых кошек с чувствительным пищеварением, 7 кг",
         oid="5543934520",
         url="https://www.ozon.ru/product/5543934520/",
         claims_short="30% свежего мяса; 75% мясного белка; легкоусвояемый белок",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Goodwin",
         name="Super Premium, для взрослых кошек, индейка и утка, 1,5 кг",
         oid="5543935300",
         url="https://www.ozon.ru/product/5543935300/",
         claims_short="Поддержка иммунной системы (розмарин, алоэ, зелёный чай); 51%; Super Premium",
         protein="Индейка, Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AlphaPet",
         name="Для кошек и котов крупных пород, 2 кг",
         oid="5543935895",
         url="https://www.ozon.ru/product/5543935895/",
         claims_short="30% свежего мяса; 75% мясного белка; легкоусвояемый белок",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Goodwin",
         name="Super Premium, для котят, индейка, 400 г",
         oid="5543936157",
         url="https://www.ozon.ru/product/5543936157/",
         claims_short="Поддержка иммунной системы (розмарин, алоэ, зелёный чай); Super Premium",
         protein="Индейка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AlphaPet",
         name="Для взрослых стерилизованных кошек и котов, с цыплёнком, 7 кг",
         oid="5543936716",
         url="https://www.ozon.ru/product/5543936716/",
         claims_short="30% свежего мяса; 75% мясного белка; легкоусвояемый белок",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Primordial",
         name="Cat Adult, утка и индейка, беззерновой, 2 кг",
         oid="5554107233",
         url="https://www.ozon.ru/product/5554107233/",
         claims_short="Grain Free; Holistic Nutrition; 100% Natural; 70% Animal Ingredients; Natural Instinct",
         protein="Утка, Индейка", ingredients="—", plant="Италия"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Primordial",
         name="Cat Adult Neutered Exigent, лосось и перепёлка, беззерновой, 2 кг",
         oid="5554109198",
         url="https://www.ozon.ru/product/5554109198/",
         claims_short="Grain Free; Holistic Nutrition; 100% Natural; Animal Ingredients; Natural Instinct",
         protein="Лосось, Перепел", ingredients="—", plant="Италия"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Primordial",
         name="Cat Adult Neutered Skin&Coat, лосось и индейка, беззерновой, 6 кг",
         oid="5554109696",
         url="https://www.ozon.ru/product/5554109696/",
         claims_short="Grain Free; Holistic Nutrition; 100% Natural; Animal Ingredients; здоровье кожи и шерсти; Natural Instinct",
         protein="Лосось, Индейка", ingredients="—", plant="Италия"),
    dict(animal="Кошки", ptype="Сухой корм", brand="TASTY",
         name="С говядиной, для взрослых кошек, 10 кг",
         oid="5555801791",
         url="https://www.ozon.ru/product/5555801791/",
         claims_short="Без заявленных клеймов — базовое описание",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Monge",
         name="Cat Monoprotein Sterilised, утка, 1,5 кг",
         oid="5556095810",
         url="https://www.ozon.ru/product/5556095810/",
         claims_short="Monoprotein; First Ingredient №1 Duck; enhanced palatability; sub-caloric content; Natural Superpremium",
         protein="Утка", ingredients="—", plant="Италия"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Landor",
         name="Sterilised Adult, индейка с уткой, 10 кг",
         oid="5556862888",
         url="https://www.ozon.ru/product/5556862888/",
         claims_short="Super Premium Food Class; не содержит пшеницы; без ГМО",
         protein="Индейка, Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AlphaPet",
         name="Superpremium Monoprotein, для котят, беременных и кормящих кошек, с индейкой, 400 г",
         oid="5566319732",
         url="https://www.ozon.ru/product/5566319732/",
         claims_short="30% свежего мяса; 75% мясного белка; легкоусвояемый белок; Superpremium",
         protein="Индейка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AlphaPet",
         name="Superpremium, для пожилых кошек от 7 лет, кролик и утка, 400 г",
         oid="5566555324",
         url="https://www.ozon.ru/product/5566555324/",
         claims_short="30% свежего мяса; 75% мясного белка; легкоусвояемый белок; Superpremium",
         protein="Кролик, Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AlphaPet",
         name="Superpremium, с говядиной и печенью, 0,4 кг",
         oid="5566602205",
         url="https://www.ozon.ru/product/5566602205/",
         claims_short="30% свежего мяса; 75% мясного белка; легкоусвояемый белок; Superpremium",
         protein="Говядина, Печень", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AlphaPet Menu",
         name="Adult, с говядиной, свежее мясо, для взрослых кошек и котов, 1,5 кг",
         oid="5566632364",
         url="https://www.ozon.ru/product/5566632364/",
         claims_short="Со свежим мясом; 100% натуральные ингредиенты; сбалансированный рацион",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="AlphaPet",
         name="Superpremium, для взрослых кошек, с цыплёнком, 1,5 кг",
         oid="5566657335",
         url="https://www.ozon.ru/product/5566657335/",
         claims_short="30% свежего мяса; 75% мясного белка; легкоусвояемый белок; Superpremium",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Karmy",
         name="Hypoallergenic, для кошек, склонных к аллергии, утка, 10 кг",
         oid="5568745202",
         url="https://www.ozon.ru/product/5568745202/",
         claims_short="Гипоаллергенный; высокое содержание мяса; улучшает работу пищеварительной системы; поддерживает здоровье кожи и шерсти",
         protein="Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Optima Nova",
         name="Беззерновой, с кроликом и картофелем, для взрослых кошек, 8 кг",
         oid="5568879450",
         url="https://www.ozon.ru/product/5568879450/",
         claims_short="Grain Free; 75% свежего кролика (+30% дегидрированного, 5% гидролизата); Omega 3&6; здоровье кожи и шерсти; высокая усвояемость",
         protein="Кролик", ingredients="Картофель; Омега-3/6", plant="Испания"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Purina Pro Plan",
         name="Veterinary Diets EN ST/OX Gastrointestinal, для взрослых кошек и котят при кишечных расстройствах, 1,5 кг х 4",
         oid="5574288140",
         url="https://www.ozon.ru/product/5574288140/",
         claims_short="VET; ветеринарная диета для снижения проявлений кишечных расстройств",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Сухой корм", brand="Hilton",
         name="Adult Sensitive, с чувствительным пищеварением, кролик, 2х2 кг",
         oid="5578339795",
         url="https://www.ozon.ru/product/5578339795/",
         claims_short="Гипоаллергенный корм; сбалансированный рацион; комплекс витаминов; Eco-Premium; Super Premium",
         protein="Кролик", ingredients="—", plant="Не найдено"),

    # --- Day 7 (2026-08-24): Влажный корм для кошек ---
    dict(animal="Кошки", ptype="Влажный корм", brand="Gourmet Gold",
         name="Паштет с курицей, 85 г х 6",
         oid="5484409205",
         url="https://www.ozon.ru/product/5484409205/",
         claims_short="Произведено во Франции",
         protein="Курица", ingredients="—", plant="Франция"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Gourmet Gold",
         name="Паштет с индейкой, 85 г х 6",
         oid="5484413359",
         url="https://www.ozon.ru/product/5484413359/",
         claims_short="Произведено во Франции",
         protein="Индейка", ingredients="—", plant="Франция"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Gourmet Gold",
         name="Паштет с тунцом, 85 г х 6",
         oid="5484418520",
         url="https://www.ozon.ru/product/5484418520/",
         claims_short="Произведено во Франции",
         protein="Тунец", ingredients="—", plant="Франция"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Wanpy Cat",
         name="Тунец в собственном соку, пауч, 70 г х 12",
         oid="5530649635",
         url="https://www.ozon.ru/product/5530649635/",
         claims_short="Real Meat; No Preservatives, No Colorants; Healthy Digestion, Easy to Absorb; Super Premium",
         protein="Тунец", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Wanpy Cat",
         name="Кусочки курицы и креветки, пауч, 85 г х 12",
         oid="5530649765",
         url="https://www.ozon.ru/product/5530649765/",
         claims_short="Real Meat; No Preservatives, No Colorants; Healthy Digestion, Easy to Absorb; Super Premium",
         protein="Курица, Креветки", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Wanpy Cat",
         name="Кусочки курицы и лосось, 95 г х 12",
         oid="5530650056",
         url="https://www.ozon.ru/product/5530650056/",
         claims_short="Real Meat; No Preservatives, No Colorants; Healthy Digestion, Easy to Absorb; Super Premium",
         protein="Курица, Лосось", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Wanpy Cat",
         name="Кусочки курицы и гребешки, пауч, 85 г х 12",
         oid="5530367150",
         url="https://www.ozon.ru/product/5530367150/",
         claims_short="Real Meat; No Preservatives, No Colorants; Healthy Digestion, Easy to Absorb; Super Premium",
         protein="Курица, Гребешок", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Wanpy Cat",
         name="Нежное рагу из лосося и тунца, для котят, пауч, 70 г х 12",
         oid="5530650988",
         url="https://www.ozon.ru/product/5530650988/",
         claims_short="Real Meat; No Preservatives, No Colorants; Healthy Digestion, Easy to Absorb; Super Premium",
         protein="Лосось, Тунец", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Wanpy Cat",
         name="Нежное рагу из курицы и утки, для котят, пауч, 70 г х 12",
         oid="5530651154",
         url="https://www.ozon.ru/product/5530651154/",
         claims_short="Real Meat; No Preservatives, No Colorants; Healthy Digestion, Easy to Absorb; Super Premium",
         protein="Курица, Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Wanpy Cat",
         name="Кусочки курицы и треска, 95 г х 12",
         oid="5530651642",
         url="https://www.ozon.ru/product/5530651642/",
         claims_short="Real Meat; No Preservatives, No Colorants; Healthy Digestion, Easy to Absorb; Super Premium",
         protein="Курица, Треска", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Wanpy Cat",
         name="Кусочки курицы и тунец, 95 г х 12",
         oid="5530653845",
         url="https://www.ozon.ru/product/5530653845/",
         claims_short="Real Meat; No Preservatives, No Colorants; Healthy Digestion, Easy to Absorb; Super Premium",
         protein="Курица, Тунец", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Wanpy Cat",
         name="Кусочки курицы и краб, пауч, 85 г х 12",
         oid="5530654016",
         url="https://www.ozon.ru/product/5530654016/",
         claims_short="Real Meat; No Preservatives, No Colorants; Healthy Digestion, Easy to Absorb; Super Premium",
         protein="Курица, Краб", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="—",
         name="Индейка и курица, 85 г х 24",
         oid="5538615147",
         url="https://www.ozon.ru/product/5538615147/",
         claims_short="Без заявленных клеймов — базовое описание",
         protein="Индейка, Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Мнямс",
         name="Деликатес, ассорти двухслойных паштетов из курицы (говядина/кролик/утка), для взрослых и стерилизованных кошек, 6х40 г",
         oid="5543025974",
         url="https://www.ozon.ru/product/5543025974/",
         claims_short="Без заявленных клеймов — базовое описание",
         protein="Курица, Говядина, Кролик, Утка", ingredients="Морковь; Кошачья мята; Клюква", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="—",
         name="С говядиной, 100 г",
         oid="5544350676",
         url="https://www.ozon.ru/product/5544350676/",
         claims_short="Без заявленных клеймов — базовое описание",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="—",
         name="С мясным ассорти, 100 г",
         oid="5544357753",
         url="https://www.ozon.ru/product/5544357753/",
         claims_short="Без заявленных клеймов — базовое описание",
         protein="Не конкретизирован на лицевой стороне упаковки (мясное ассорти)", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Royal Canin",
         name="Kitten, паштет в соусе, для котят до 12 месяцев, 3 пауча по 85 г",
         oid="5551551944",
         url="https://www.ozon.ru/product/5551551944/",
         claims_short="No Preservatives; Health Nutrition since 1968",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Schesir",
         name="After Dark, паштет из курицы и перепелиного яйца, беззерновой, консервы, 12х80 г",
         oid="5566574938",
         url="https://www.ozon.ru/product/5566574938/",
         claims_short="Grain Free; Complete",
         protein="Курица, Перепелиное яйцо", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Schesir",
         name="After Dark, паштет из курицы, беззерновой, консервы, 12х80 г",
         oid="5566598587",
         url="https://www.ozon.ru/product/5566598587/",
         claims_short="Grain Free; Complete",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Schesir",
         name="After Dark, паштет из курицы и говядины, беззерновой, консервы, 12х80 г",
         oid="5566620215",
         url="https://www.ozon.ru/product/5566620215/",
         claims_short="Grain Free; Complete",
         protein="Курица, Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="AlphaPet WOW",
         name="Superpremium, кролик с потрошками, кусочки в соусе, для чувствительного пищеварения, 28х85 г",
         oid="5654017462",
         url="https://www.ozon.ru/product/5654017462/",
         claims_short="Чувствительное пищеварение; здоровое пищеварение; Fresh Meat (Rich Meat Profile); Superpremium; комплекс AlphaPetBIO",
         protein="Кролик", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="P.E.P.P.O.",
         name="Курица, кусочки в соусе, 24х85 г",
         oid="5664781811",
         url="https://www.ozon.ru/product/5664781811/",
         claims_short="Натуральные ингредиенты; поддержка иммунитета; 85% мяса в каждом кусочке",
         protein="Курица", ingredients="Фукус; Юкка; Черника", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="P.E.P.P.O.",
         name="Ягнёнок, кусочки в соусе, для котят, 24х85 г",
         oid="5664784502",
         url="https://www.ozon.ru/product/5664784502/",
         claims_short="Натуральные ингредиенты; поддержка иммунитета; 85% мяса в каждом кусочке",
         protein="Ягнёнок", ingredients="Фукус; Юкка; Черника", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Вкусная Миска",
         name="Говядина в нежном соусе, дополнительный корм, 5х85 г",
         oid="5668947557",
         url="https://www.ozon.ru/product/5668947557/",
         claims_short="С витаминами и минералами; без красителей; без ГМО",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Вкусная Миска",
         name="Кролик в соусе, дополнительный корм, 5х85 г",
         oid="5669095176",
         url="https://www.ozon.ru/product/5669095176/",
         claims_short="С витаминами и минералами; без красителей; без ГМО",
         protein="Кролик", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Whiskas",
         name="Особое удовольствие, желе с индейкой, печенью и овощами, 28х75 г",
         oid="5670228940",
         url="https://www.ozon.ru/product/5670228940/",
         claims_short="Особое удовольствие (вкусовая привлекательность)",
         protein="Индейка, Печень", ingredients="Овощи", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="SHEBA",
         name="Ломтики в соусе, кролик и индейка, 28х75 г",
         oid="5670237889",
         url="https://www.ozon.ru/product/5670237889/",
         claims_short="Ломтики в соусе (вкусовая привлекательность)",
         protein="Кролик, Индейка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="SHEBA",
         name="Ломтики в соусе, курица, 28х75 г",
         oid="5670251820",
         url="https://www.ozon.ru/product/5670251820/",
         claims_short="Ломтики в соусе (вкусовая привлекательность)",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Perfect Fit",
         name="С чувствительным пищеварением, индейка в соусе, 28х75 г",
         oid="5670328613",
         url="https://www.ozon.ru/product/5670328613/",
         claims_short="Чувствительное пищеварение; здоровая микрофлора кишечника; без добавления пшеницы и сои; разработано со специалистами Waltham Petcare Science Institute",
         protein="Индейка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Royal Canin",
         name="Sterilised, консервированный корм в соусе, для стерилизованных взрослых кошек, 85 г",
         oid="5670392893",
         url="https://www.ozon.ru/product/5670392893/",
         claims_short="Health Nutrition since 1968",
         protein="Не указан на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Мнямс (Ферма кота Фёдора)",
         name="Сочные кусочки с лососем, креветками и форелью, пауч, 85 г х 8",
         oid="5670714887",
         url="https://www.ozon.ru/product/5670714887/",
         claims_short="Без ГМО; без сои; витамин Е",
         protein="Лосось, Креветки, Форель", ingredients="Витамин Е", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Мнямс (Ферма кота Фёдора)",
         name="Нежные кусочки с индейкой в желе, для котят, пауч, 85 г",
         oid="5670716904",
         url="https://www.ozon.ru/product/5670716904/",
         claims_short="Без ГМО; без сои",
         protein="Индейка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Мнямс (Ферма кота Фёдора)",
         name="Нежные кусочки с телятиной, для котят, пауч, 85 г х 16",
         oid="5670719609",
         url="https://www.ozon.ru/product/5670719609/",
         claims_short="Без ГМО; без сои",
         protein="Телятина", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Мнямс (Ферма кота Фёдора)",
         name="Сочные кусочки с говядиной в желе, пауч, 85 г х 8",
         oid="5670721156",
         url="https://www.ozon.ru/product/5670721156/",
         claims_short="Без ГМО; без сои",
         protein="Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Мнямс",
         name="Кусочки в соусе с уткой, для домашних кошек, пауч, 85 г х 6",
         oid="5671557401",
         url="https://www.ozon.ru/product/5671557401/",
         claims_short="Профилактика мочекаменной болезни; поддержание оптимальной кондиции",
         protein="Утка", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Влажный корм", brand="Schesir",
         name="After Dark, паштет из курицы и утки, беззерновой, консервы, 12х80 г",
         oid="5566642607",
         url="https://www.ozon.ru/product/5566642607/",
         claims_short="Grain Free; Complete",
         protein="Курица, Утка", ingredients="—", plant="Не найдено"),

    # --- Day 7 (2026-08-24): Лакомства для кошек ---
    # 5577401282 (говяжья трахея 100г) — тот же товар/линейка, что уже учтённый
    # 5577402177 у собак (дубль по фасовке, ошибочно всплыл в выдаче категории кошек);
    # 5557001158 — пустой/удалённый листинг без названия и фото, не профилируется.
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Paste, куриное филе с лактобактериями, для здоровья кожи и шерсти, 80 г",
         oid="5632614370",
         url="https://www.ozon.ru/product/5632614370/",
         claims_short="Здоровье кожи и шерсти",
         protein="Курица", ingredients="Лактобактерии", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Churu, лакомство-пюре, тунец Кацуо, 14 г х 4",
         oid="5632614943",
         url="https://www.ozon.ru/product/5632614943/",
         claims_short="Вкусовая привлекательность",
         protein="Тунец", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="Yaki Sasami, куриное филе со вкусом краба, 26 г",
         oid="5632615075",
         url="https://www.ozon.ru/product/5632615075/",
         claims_short="Вкусовая привлекательность",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="Yaki Miksu, запечённое ассорти, кацуобуси, куриный бульон и кальмар, 25 г",
         oid="5632615106",
         url="https://www.ozon.ru/product/5632615106/",
         claims_short="Вкусовая привлекательность",
         protein="Курица, Кальмар", ingredients="Кацуобуси", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Churu, лакомство-пюре, куриное филе, для здоровья кожи и шерсти, 14 г х 4",
         oid="5632619730",
         url="https://www.ozon.ru/product/5632619730/",
         claims_short="Здоровье кожи и шерсти",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Churu, лакомство-пюре, тунец Магуро, 14 г х 4",
         oid="5654586643",
         url="https://www.ozon.ru/product/5654586643/",
         claims_short="Вкусовая привлекательность",
         protein="Тунец", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="Tobitsuku, запечённое лакомство, со вкусом куриного бульона, 25 г",
         oid="5654587557",
         url="https://www.ozon.ru/product/5654587557/",
         claims_short="Вкусовая привлекательность",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Churu, лакомство-пюре, куриное филе, поддержание pH мочи, 14 г х 4",
         oid="5654590710",
         url="https://www.ozon.ru/product/5654590710/",
         claims_short="Поддержание pH мочи",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Pon Churu, суфле-паштет, тунец Магуро, 35 г х 2",
         oid="5654591217",
         url="https://www.ozon.ru/product/5654591217/",
         claims_short="Вкусовая привлекательность",
         protein="Тунец", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Churu, лакомство-пюре, куриное филе и краб, 14 г х 4",
         oid="5654591653",
         url="https://www.ozon.ru/product/5654591653/",
         claims_short="Вкусовая привлекательность",
         protein="Курица, Краб", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Churu, лакомство-пюре, тунец Магуро и морской гребешок, 14 г х 4",
         oid="5654592626",
         url="https://www.ozon.ru/product/5654592626/",
         claims_short="Вкусовая привлекательность",
         protein="Тунец, Гребешок", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Churu, лакомство-пюре, тунец Магуро и лобстер, 14 г х 4",
         oid="5654594462",
         url="https://www.ozon.ru/product/5654594462/",
         claims_short="Вкусовая привлекательность",
         protein="Тунец, Лобстер", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Churu, лакомство-пюре, куриное филе, для здоровья мочеполовой системы, 14 г х 4",
         oid="5654596008",
         url="https://www.ozon.ru/product/5654596008/",
         claims_short="Здоровье мочеполовой системы",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Churu, лакомство-пюре, куриное филе и кальмар, 14 г х 4",
         oid="5654598009",
         url="https://www.ozon.ru/product/5654598009/",
         claims_short="Вкусовая привлекательность",
         protein="Курица, Кальмар", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="Tobitsuku, запечённое лакомство, со вкусом морского гребешка, 25 г",
         oid="5654598442",
         url="https://www.ozon.ru/product/5654598442/",
         claims_short="Вкусовая привлекательность",
         protein="Гребешок", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Churu, лакомство-пюре, куриное филе, для котят, 14 г х 4",
         oid="5654598890",
         url="https://www.ozon.ru/product/5654598890/",
         claims_short="Вкусовая привлекательность",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="Yaki Sasami, куриное филе со вкусом говядины, 26 г",
         oid="5654599052",
         url="https://www.ozon.ru/product/5654599052/",
         claims_short="Вкусовая привлекательность",
         protein="Курица, Говядина", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Pon Churu, суфле-паштет, куриное филе и морской гребешок, 35 г х 2",
         oid="5654599091",
         url="https://www.ozon.ru/product/5654599091/",
         claims_short="Вкусовая привлекательность",
         protein="Курица, Гребешок", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="INABA",
         name="CIAO Churu, лакомство-пюре, куриное филе и креветка, 14 г х 4",
         oid="5654599124",
         url="https://www.ozon.ru/product/5654599124/",
         claims_short="Вкусовая привлекательность",
         protein="Курица, Креветки", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Мнямс",
         name="Крем-лакомство, утка с тунцом Кацуо, 4х15 г",
         oid="5670711038",
         url="https://www.ozon.ru/product/5670711038/",
         claims_short="Нежное пюре-лакомство",
         protein="Утка, Тунец", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Мнямс",
         name="Крем-лакомство, курица, 4х15 г",
         oid="5671503941",
         url="https://www.ozon.ru/product/5671503941/",
         claims_short="Нежное пюре-лакомство",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="АШАН Pets",
         name="Хрустящие сердечки, индейка со сливками, 40 г х 3",
         oid="5542724663",
         url="https://www.ozon.ru/product/5542724663/",
         claims_short="Без добавления искусственных ароматизаторов; витамины и минералы в составе; подходит для стерилизованных кошек и котят",
         protein="Индейка", ingredients="Сливки", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Ферма кота Федора",
         name="Лакомые палочки, кролик, 3х5 г",
         oid="5543308856",
         url="https://www.ozon.ru/product/5543308856/",
         claims_short="Без ГМО; без сои; таурин",
         protein="Кролик", ingredients="Таурин", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Ферма кота Федора",
         name="Лакомые палочки, лосось, 3х5 г",
         oid="5543534231",
         url="https://www.ozon.ru/product/5543534231/",
         claims_short="Без ГМО; без сои; таурин",
         protein="Лосось", ingredients="Таурин", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Ферма кота Федора",
         name="Лакомые палочки, утка, 3х5 г",
         oid="5543551335",
         url="https://www.ozon.ru/product/5543551335/",
         claims_short="Без ГМО; без сои; таурин",
         protein="Утка", ingredients="Таурин", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Dreamies",
         name="Молочный коктейль, подушечки с добавлением молока, 8х60 г",
         oid="5549365822",
         url="https://www.ozon.ru/product/5549365822/",
         claims_short="Без искусственных ароматизаторов; витамины и минералы; высокое содержание белка",
         protein="Не конкретизирован на лицевой стороне упаковки", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="AWARD",
         name="Special Care, крем-лакомство, утка с тунцом Кацуо, чувствительное пищеварение, 4х15 г",
         oid="5579052755",
         url="https://www.ozon.ru/product/5579052755/",
         claims_short="100% BIO; Fresh Meat; Nutraceutical Formula; Quality for Champions; чувствительное пищеварение",
         protein="Утка, Тунец", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="AWARD",
         name="Special Care, крем-лакомство, курица, профилактика МКБ, 4х15 г",
         oid="5579058668",
         url="https://www.ozon.ru/product/5579058668/",
         claims_short="100% BIO; Fresh Meat; Nutraceutical Formula; Quality for Champions; профилактика мочекаменной болезни",
         protein="Курица", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="AWARD",
         name="Special Care, крем-лакомство, лосось с тунцом Кацуо, здоровая кожа и шерсть, 4х15 г",
         oid="5579066625",
         url="https://www.ozon.ru/product/5579066625/",
         claims_short="100% BIO; Fresh Meat; Nutraceutical Formula; Quality for Champions; здоровая кожа и шерсть",
         protein="Лосось, Тунец", ingredients="—", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Finesse Licko",
         name="Крем-лакомство в стиках, курица с тыквенной клетчаткой, for digestive care, 5х14 г",
         oid="5582640835",
         url="https://www.ozon.ru/product/5582640835/",
         claims_short="Без искусственных красителей; удобно давать с руки; for digestive care",
         protein="Курица", ingredients="Тыквенная клетчатка", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Murkel",
         name="Крем-пюре, курица, рыбий жир, козье молоко, 4х15 г",
         oid="5584169750",
         url="https://www.ozon.ru/product/5584169750/",
         claims_short="Натуральный состав; нежная кремовая текстура; для поддержания пищеварения/иммунитета",
         protein="Курица", ingredients="Рыбий жир; Козье молоко", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Murkel",
         name="Крем-пюре, льняное семя, курица, печень, 4х15 г",
         oid="5584218651",
         url="https://www.ozon.ru/product/5584218651/",
         claims_short="Натуральный состав; нежная кремовая текстура; для поддержания пищеварения/иммунитета",
         protein="Курица, Печень", ingredients="Льняное семя", plant="Не найдено"),
    dict(animal="Кошки", ptype="Лакомства", brand="Murkel",
         name="Крем-пюре, тунец, молочные продукты, 4х15 г",
         oid="5584264918",
         url="https://www.ozon.ru/product/5584264918/",
         claims_short="Натуральный состав; нежная кремовая текстура; для поддержания пищеварения/иммунитета",
         protein="Тунец", ingredients="Молочные продукты", plant="Не найдено"),
]

# категория-код на товар (для подсчёта количества категорий на товар)
cats_per_product = [3, 6, 6, 4, 1, 4, 2, 5, 4, 3, 2,  3, 3, 3, 0,
                     2, 4, 3, 2, 2, 3, 0, 1, 5, 2, 3, 2,
                     # доразбор 01.08 + 03.08 (в порядке добавления в products)
                     1, 1, 1, 3, 3, 3, 3, 1, 1,
                     3, 5, 6, 1, 4, 0, 4, 1, 1, 3, 3, 3, 3, 2, 3,
                     # Day 5 (2026-08-04)
                     5, 2, 1, 2, 2, 2, 2, 2, 6, 2, 2, 3, 2, 1, 1, 1, 2,
                     # Day 6 (2026-08-12) — в порядке добавления в products
                     2, 2, 2, 2, 1, 1, 2, 1, 1, 2, 4, 1, 2, 1,
                     3, 1, 2, 2, 2, 3, 3, 3, 1, 3,
                     3, 4, 3, 3, 3, 2, 3,
                     # Доразбор 12.08 (пропущенные, найдены аудитом ?sorting=new)
                     1, 2, 2, 2, 2, 2, 0, 2, 4, 3, 2,
                     # Day 7 (2026-08-24) — Каша для собак
                     2, 2, 2, 3, 3, 3, 3, 2, 2,
                     # Day 7 (2026-08-24) — Сухой корм для собак
                     3, 3, 3, 3, 3, 3, 5, 4, 6, 2, 4, 1, 1,
                     # Day 7 (2026-08-24) — Влажный корм для собак
                     1, 2, 2, 2, 1, 1, 0, 0,
                     # Day 7 (2026-08-24) — Лакомства для собак
                     2, 2, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 2, 1, 2,
                     # Day 7 (2026-08-24) — Сухой корм для кошек
                     3, 3, 5, 2, 2, 3, 2, 2, 2, 4, 4, 5, 0, 4, 2, 2, 2, 2, 2, 2, 3, 3, 1, 3,
                     # Day 7 (2026-08-24) — Влажный корм для кошек
                     1, 1, 1, 4, 4, 4, 4, 4, 4, 4, 4, 4, 0, 0, 0, 0, 2, 1, 1, 1, 1,
                     # Day 7 (2026-08-24) — Лакомства для кошек
                     2, 1, 1, 1, 2, 4, 4, 4, 2, 2, 2, 2,
] + [0] * 74
# Day 8 (2026-09-02) и далее: значения-плейсхолдеры выше не используются —
# реальное "Кол-во категорий клеймов" пересчитывается из claim_rows после его
# определения (см. ниже, после заполнения листа "Клеймы"). Плейсхолдер нужен
# только чтобы zip(products, cats_per_product) не обрезал новые товары по
# длине списка при добавлении очередной партии.

METHOD_BY_DATE = {
    "2026-07-31": "Сортировка «Новинки» (не подтверждено ID)",
    "2026-08-01": "Подтверждено: ID выше макс. известного на 2026-07-31",
    "2026-08-03": "Подтверждено: ID выше макс. известного на 2026-08-01",
    "2026-08-04": "Подтверждено: ID выше макс. известного на 2026-08-03",
    "2026-08-12": "Подтверждено: ID выше макс. известного на 2026-08-04",
    "2026-08-24": "Подтверждено: ID выше макс. известного на 2026-08-12",
    "2026-09-02": "Подтверждено: ID выше макс. известного на 2026-08-24",
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

    # Встроенное превью фото упаковки (не ссылка, а сама картинка) — реально
    # уменьшаем и пережимаем пиксели перед вставкой, иначе openpyxl встраивает
    # исходный файл целиком (несколько сотен КБ каждый) и .xlsx быстро упирается
    # в лимит GitHub на размер файла при росте выборки.
    ppath = photo_path(p["oid"])
    im = PILImage.open(ppath)
    w0, h0 = im.size
    target_w = int(w0 * TARGET_H_PX / h0)
    max_photo_w = max(max_photo_w, target_w)
    thumb = im.convert("RGB").resize((target_w, TARGET_H_PX), PILImage.LANCZOS)
    buf = io.BytesIO()
    thumb.save(buf, format="JPEG", quality=72, optimize=True)
    buf.seek(0)
    img = XLImage(buf)
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

    # --- Day 5 (2026-08-04), подтверждено сравнением ID ---
    ("5345830359", "Pets Brunch Adult Super Mini Breeds (индейка)", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Sensitive; гипоаллергенный"),
    ("5345830359", "Pets Brunch Adult Super Mini Breeds (индейка)", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Укрепляет иммунитет; защита пищеварения; красота кожи и шерсти"),
    ("5345830359", "Pets Brunch Adult Super Mini Breeds (индейка)", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Gluten Free"),
    ("5345830359", "Pets Brunch Adult Super Mini Breeds (индейка)", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "41% мясных ингредиентов"),
    ("5345830359", "Pets Brunch Adult Super Mini Breeds (индейка)", "Собаки", "Сухой корм", "ORIGIN", "Страна / регион производства", "Belgian recipe"),

    ("5358300410", "Mil.Dog для собак мелких пород", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "Без искусственных ароматизаторов, усилителей вкуса и сои"),
    ("5358300410", "Mil.Dog для собак мелких пород", "Собаки", "Сухой корм", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5357288481", "Farmina N&D Quail & Pumpkin", "Собаки", "Влажный корм", "ORIGIN", "Страна / регион производства", "Italian pet food (флаг Италии на этикетке)"),

    ("5358242931", "Родные корма — скоблянка мясная по-городецки", "Собаки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Премиум"),
    ("5358242931", "Родные корма — скоблянка мясная по-городецки", "Собаки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Мясо ингредиент №1; 75% мяса и мясных ингредиентов"),

    ("5358240926", "Родные корма — по-восточному, баранина с потрошками", "Собаки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Премиум"),
    ("5358240926", "Родные корма — по-восточному, баранина с потрошками", "Собаки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Мясо ингредиент №1; 75% мяса и мясных ингредиентов"),

    ("5359892602", "Puffins индейка и лосось (стерилизованные)", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Low Grain"),
    ("5359892602", "Puffins индейка и лосось (стерилизованные)", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Здоровый иммунитет; контроль веса; здоровье мочевыдел. системы; здоровое пищеварение; здоровье шерсти и кожи"),

    ("5359453152", "Puffins говядина и лосось", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Low Grain"),
    ("5359453152", "Puffins говядина и лосось", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Здоровый иммунитет; высокая усвояемость пищи; здоровое пищеварение; поддержание мышечного тонуса; здоровье кожи и шерсти"),

    ("5359660535", "Puffins мясное ассорти", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Low Grain"),
    ("5359660535", "Puffins мясное ассорти", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Здоровый иммунитет; здоровье сердца и глаз; здоровое пищеварение; сильные мышцы и связки; здоровье шерсти и кожи"),

    ("5359822345", "GRANDORF Lamb & Turkey Adult Indoor (кошки)", "Кошки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic; ingredients for sensitive skin & stomach"),
    ("5359822345", "GRANDORF Lamb & Turkey Adult Indoor (кошки)", "Кошки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
    ("5359822345", "GRANDORF Lamb & Turkey Adult Indoor (кошки)", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Single Grain formula; Gluten Free"),
    ("5359822345", "GRANDORF Lamb & Turkey Adult Indoor (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "70% premium animal ingredients"),
    ("5359822345", "GRANDORF Lamb & Turkey Adult Indoor (кошки)", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "Without the addition of plant protein; No artificial flavours/colours/preservatives"),
    ("5359822345", "GRANDORF Lamb & Turkey Adult Indoor (кошки)", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Optimal ratio of Omega 3-6 fatty acids; vitamins, minerals & natural antioxidants"),

    ("5357563517", "DELISH super premium — суп с курицей", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Super Premium; премиальный состав"),
    ("5357563517", "DELISH super premium — суп с курицей", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Высокое содержание чистого мяса"),

    ("5357853793", "DELISH super premium — суп с говядиной", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),
    ("5357853793", "DELISH super premium — суп с говядиной", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Высокое содержание чистого мяса"),

    ("5331257302", "Blitz Sterilised — индейка и клюква (паучи)", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Urinary care"),
    ("5331257302", "Blitz Sterilised — индейка и клюква (паучи)", "Кошки", "Влажный корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Sensitive"),
    ("5331257302", "Blitz Sterilised — индейка и клюква (паучи)", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Высокое качество"),

    ("5358325770", "Mr.Whiskers крем-лакомство, ассорти 7 вкусов", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Лакомство для твоего пушистика; вкусное и полезное"),
    ("5358325770", "Mr.Whiskers крем-лакомство, ассорти 7 вкусов", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Improve hair health (Codfish); improve intestinal health (Salmon); improve urinary tract (Rabbit); improve immunity (Tuna)"),

    ("5358242441", "Мнямс — Ферма кота Федора, тунец", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Без ГМО; без сои"),
    ("5358240663", "Мнямс — Ферма кота Федора, говядина", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Без ГМО; без сои"),
    ("5358241683", "Мнямс — Ферма кота Федора, курица", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Без ГМО; без сои"),

    ("5358240549", "Четыре с хвостиком — травка для кошек", "Кошки", "Лакомства", "NAT", "Натуральность состава", "100% натуральный продукт"),
    ("5358240549", "Четыре с хвостиком — травка для кошек", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Для выведения шерсти; спасает комнатные растения"),

    # --- Day 6 (2026-08-12) ---
    ("5430117217", "АЛЬМАКОРМ Каша с говядиной", "Собаки", "Каши", "NAT", "Натуральность состава", "100% натуральный, без вкусовых добавок"),
    ("5430117217", "АЛЬМАКОРМ Каша с говядиной", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки; 5 минут и готово"),

    ("5423933879", "АЛЬМАКОРМ Каша с индейкой", "Собаки", "Каши", "NAT", "Натуральность состава", "100% натуральный, без вкусовых добавок"),
    ("5423933879", "АЛЬМАКОРМ Каша с индейкой", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки; 5 минут и готово"),

    ("5423741999", "АЛЬМАКОРМ Каша свинина с говядиной", "Собаки", "Каши", "NAT", "Натуральность состава", "100% натуральный, без вкусовых добавок"),
    ("5423741999", "АЛЬМАКОРМ Каша свинина с говядиной", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки; 5 минут и готово"),

    ("5432594014", "СибКорм Каша с рыбой", "Собаки", "Каши", "NAT", "Натуральность состава", "Без ГМО; из натуральных ингредиентов"),
    ("5432594014", "СибКорм Каша с рыбой", "Собаки", "Каши", "ORIGIN", "Страна / регион производства", "Произведено на Алтае"),

    ("5419467431", "Limuvil Каша для собак", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки"),

    ("5418435027", "Командирская Каша для собак", "Собаки", "Каши", "ORIGIN", "Страна / регион производства", "С Алтайских полей"),

    ("5417460160", "Альма Каша с говядиной", "Собаки", "Каши", "NAT", "Натуральность состава", "Без вкусовых добавок"),
    ("5417460160", "Альма Каша с говядиной", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки; 5 минут и готово"),

    ("5432515655", "Justin курица с печенью в соусе", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Витамины и минералы; таурин; протеин"),

    ("5431773123", "Мнямс микс кролик/лосось (стерилизованные)", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Профилактика мочекаменной болезни; контроль веса"),

    ("5384510562", "Blitz Sterilised утка с лососем", "Кошки", "Влажный корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
    ("5384510562", "Blitz Sterilised утка с лососем", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "90% мясных ингредиентов в каждом кусочке"),

    ("5429960930", "Animonda Carny курица+индейка+утка", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "100% свежие мясные ингредиенты; высокое содержание животных протеинов"),
    ("5429960930", "Animonda Carny курица+индейка+утка", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Без злаков"),
    ("5429960930", "Animonda Carny курица+индейка+утка", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Без добавления сахара"),
    ("5429960930", "Animonda Carny курица+индейка+утка", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Бренд с 1993 года (seit 1993)"),

    ("5430728946", "CLAN VET Renal", "Кошки", "Влажный корм", "VET", "Ветеринарная / лечебная диета", "Профилактика болезней почек; профилактическая диета"),

    ("5430727673", "Royal Canin Light Weight Care", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Помогает контролировать вес; свыше 90% кошек достигли более здорового веса за 8 недель"),
    ("5430727673", "Royal Canin Light Weight Care", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Health Nutrition since 1968"),

    ("5430194418", "Animonda Vom Feinsten Raffinesse курица с треской", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Линейка «Raffinesse» (изысканность/утончённость)"),

    ("5424780704", "Farmina N&D Ocean лосось+треска+дыня", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5424780704", "Farmina N&D Ocean лосось+треска+дыня", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Nutrition System for Carnivores"),
    ("5424780704", "Farmina N&D Ocean лосось+треска+дыня", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Italian formula"),

    ("5424781094", "Farmina N&D Quinoa Weight Management ягнёнок", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Weight Management (контроль веса)"),

    ("5424780086", "Farmina N&D Ancestral Grain курица (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Nutrition System for Carnivores"),
    ("5424780086", "Farmina N&D Ancestral Grain курица (кошки)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Italian formula"),

    ("5424781029", "Farmina N&D Ancestral Grain Neutered курица (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Nutrition System for Carnivores"),
    ("5424781029", "Farmina N&D Ancestral Grain Neutered курица (кошки)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Italian formula"),

    ("5424783286", "Farmina N&D Ancestral Grain ягнёнок+черника (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Nutrition System for Carnivores"),
    ("5424783286", "Farmina N&D Ancestral Grain ягнёнок+черника (кошки)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Italian formula"),

    ("5424776640", "Farmina N&D Pumpkin перепел (кошки)", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5424776640", "Farmina N&D Pumpkin перепел (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Nutrition System for Carnivores"),
    ("5424776640", "Farmina N&D Pumpkin перепел (кошки)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Italian formula"),

    ("5424785573", "Farmina N&D Pumpkin кабан (собаки)", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5424785573", "Farmina N&D Pumpkin кабан (собаки)", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Nutrition System for Carnivores"),
    ("5424785573", "Farmina N&D Pumpkin кабан (собаки)", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Italian formula"),

    ("5424783287", "Farmina N&D Pumpkin Neutered ягнёнок+черника (кошки)", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5424783287", "Farmina N&D Pumpkin Neutered ягнёнок+черника (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Nutrition System for Carnivores"),
    ("5424783287", "Farmina N&D Pumpkin Neutered ягнёнок+черника (кошки)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Italian formula"),

    ("5424784781", "Farmina N&D Quinoa Neutered утка (собаки)", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Weight Management (контроль веса)"),

    ("5424782869", "Farmina N&D Prime курица (котята)", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5424782869", "Farmina N&D Prime курица (котята)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Nutrition System for Carnivores"),
    ("5424782869", "Farmina N&D Prime курица (котята)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Italian formula"),

    ("5428395063", "TiTBiT Fresh Snack (мята/говядина)", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Для освежения зубов"),
    ("5428395063", "TiTBiT Fresh Snack (мята/говядина)", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Укрепление иммунитета"),
    ("5428395063", "TiTBiT Fresh Snack (мята/говядина)", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Для дрессировки"),

    ("5428925488", "TiTBiT (Biff) Dent говядина", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Для чистки зубов; массирует дёсны"),
    ("5428925488", "TiTBiT (Biff) Dent говядина", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Укрепление иммунитета"),
    ("5428925488", "TiTBiT (Biff) Dent говядина", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Для дрессировки"),
    ("5428925488", "TiTBiT (Biff) Dent говядина", "Собаки", "Лакомства", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5429259395", "TiTBiT (Biff) Dent творог", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Для чистки зубов; массирует дёсны"),
    ("5429259395", "TiTBiT (Biff) Dent творог", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Укрепление иммунитета"),
    ("5429259395", "TiTBiT (Biff) Dent творог", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Для дрессировки"),

    ("5431102190", "TiTBiT (Biff) Dent телятина", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Для чистки зубов; массирует дёсны"),
    ("5431102190", "TiTBiT (Biff) Dent телятина", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Укрепление иммунитета"),
    ("5431102190", "TiTBiT (Biff) Dent телятина", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Для дрессировки"),

    ("5430211436", "AWARD Monoprotein паштет телятина", "Собаки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Monoprotein"),
    ("5430211436", "AWARD Monoprotein паштет телятина", "Собаки", "Влажный корм", "NAT", "Натуральность состава", "100% BIO"),
    ("5430211436", "AWARD Monoprotein паштет телятина", "Собаки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Superpremium Dog Food; Quality for Champions"),

    ("5432604231", "Karmy Starter индейка", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Способствует формированию пищеварительной системы; поддерживает иммунную систему"),
    ("5432604231", "Karmy Starter индейка", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Свежее мясо"),

    ("5432823007", "Pedigree Dentastix Light индейка", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Снижает образование зубного камня; ежедневный уход за зубами"),
    ("5432823007", "Pedigree Dentastix Light индейка", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "На 30% меньше калорий"),
    ("5432823007", "Pedigree Dentastix Light индейка", "Собаки", "Лакомства", "VET", "Ветеринарная / лечебная диета", "Разработано с ветеринарами"),

    # --- Доразбор 12.08 (найдены аудитом ?sorting=new) ---
    ("5433103013", "CLAN VET Gastrointestinal", "Собаки", "Влажный корм", "VET", "Ветеринарная / лечебная диета", "Профилактика болезней ЖКТ; профилактическая диета"),

    ("5433102619", "CLAN Classic кусочки в соусе, кролик", "Собаки", "Влажный корм", "NAT", "Натуральность состава", "Натуральные ингредиенты в составе"),
    ("5433102619", "CLAN Classic кусочки в соусе, кролик", "Собаки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Клюква — профилактика МКБ; ромашка — улучшает пищеварение"),

    ("5433103250", "Berkley №4 перепёлка с курицей и овощами", "Собаки", "Влажный корм", "NAT", "Натуральность состава", "100% Natural Product"),
    ("5433103250", "Berkley №4 перепёлка с курицей и овощами", "Собаки", "Влажный корм", "ORIGIN", "Страна / регион производства", "Local"),

    ("5433103558", "Berkley №6 кролик с говядиной и болгарским перцем", "Собаки", "Влажный корм", "NAT", "Натуральность состава", "100% Natural Product"),
    ("5433103558", "Berkley №6 кролик с говядиной и болгарским перцем", "Собаки", "Влажный корм", "ORIGIN", "Страна / регион производства", "Local"),

    ("5433104303", "Berkley №7 индейка с яблоками и отрубями", "Собаки", "Влажный корм", "NAT", "Натуральность состава", "100% Natural Product"),
    ("5433104303", "Berkley №7 индейка с яблоками и отрубями", "Собаки", "Влажный корм", "ORIGIN", "Страна / регион производства", "Local"),

    ("5433102991", "Delish консервы, кролик", "Собаки", "Влажный корм", "NAT", "Натуральность состава", "Натуральный состав; без ГМО"),
    ("5433102991", "Delish консервы, кролик", "Собаки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "100% мяса"),

    ("5433103171", "Delish консервы, говядина и ягнёнок", "Собаки", "Влажный корм", "NAT", "Натуральность состава", "Натуральный состав; без ГМО"),
    ("5433103171", "Delish консервы, говядина и ягнёнок", "Собаки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "100% мяса"),

    ("5436654518", "Кинолог в берцах — утиные головы", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Чистые зубы и здоровые дёсны"),
    ("5436654518", "Кинолог в берцах — утиные головы", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Снимает стресс; источник белка и аминокислот; легко усваивается"),
    ("5436654518", "Кинолог в берцах — утиные головы", "Собаки", "Лакомства", "NAT", "Натуральность состава", "Без красителей, консервантов и усилителей вкуса"),
    ("5436654518", "Кинолог в берцах — утиные головы", "Собаки", "Лакомства", "GRAIN", "Беззерновой / монобелковый состав", "100% утка, монопротеин"),

    ("5424786354", "Farmina N&D Prime курица и гранат (собаки)", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5424786354", "Farmina N&D Prime курица и гранат (собаки)", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Nutrition System for Carnivores"),
    ("5424786354", "Farmina N&D Prime курица и гранат (собаки)", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Italian formula"),

    ("5424780417", "Farmina N&D Ancestral Grain ягнёнок+черника (собаки)", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Nutrition System for Carnivores"),
    ("5424780417", "Farmina N&D Ancestral Grain ягнёнок+черника (собаки)", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Italian formula"),

    # --- Day 7 (2026-08-24): Каша для собак ---
    ("5542503462", "АЛЬМАКОРМ Каша с курицей", "Собаки", "Каши", "NAT", "Натуральность состава", "100% натуральный, без вкусовых добавок"),
    ("5542503462", "АЛЬМАКОРМ Каша с курицей", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки; 5 минут и готово"),

    ("5445105149", "АЛЬМАКОРМ Каша с рыбой", "Собаки", "Каши", "NAT", "Натуральность состава", "100% натуральный, без вкусовых добавок"),
    ("5445105149", "АЛЬМАКОРМ Каша с рыбой", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки; 5 минут и готово"),

    ("5445127064", "АЛЬМАКОРМ Каша с гречкой", "Собаки", "Каши", "NAT", "Натуральность состава", "100% натуральный, без вкусовых добавок"),
    ("5445127064", "АЛЬМАКОРМ Каша с гречкой", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки; 5 минут и готово"),

    ("5456129061", "ЭТИ ЗВЕРИ Лапа Злак", "Собаки", "Каши", "NAT", "Натуральность состава", "Натуральный продукт"),
    ("5456129061", "ЭТИ ЗВЕРИ Лапа Злак", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки"),
    ("5456129061", "ЭТИ ЗВЕРИ Лапа Злак", "Собаки", "Каши", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5456129661", "ЭТИ ЗВЕРИ, рисовая", "Собаки", "Каши", "NAT", "Натуральность состава", "Натуральный продукт"),
    ("5456129661", "ЭТИ ЗВЕРИ, рисовая", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки"),
    ("5456129661", "ЭТИ ЗВЕРИ, рисовая", "Собаки", "Каши", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5456131380", "ЭТИ ЗВЕРИ, злаковая", "Собаки", "Каши", "NAT", "Натуральность состава", "Натуральный продукт"),
    ("5456131380", "ЭТИ ЗВЕРИ, злаковая", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки"),
    ("5456131380", "ЭТИ ЗВЕРИ, злаковая", "Собаки", "Каши", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5456131513", "ЭТИ ЗВЕРИ, гречневая", "Собаки", "Каши", "NAT", "Натуральность состава", "Натуральный продукт"),
    ("5456131513", "ЭТИ ЗВЕРИ, гречневая", "Собаки", "Каши", "CONV", "Удобство использования", "Не требует варки"),
    ("5456131513", "ЭТИ ЗВЕРИ, гречневая", "Собаки", "Каши", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5479813617", "СибКорм Каша с мясом", "Собаки", "Каши", "NAT", "Натуральность состава", "100% натуральный состав; без ГМО"),
    ("5479813617", "СибКорм Каша с мясом", "Собаки", "Каши", "ORIGIN", "Страна / регион производства", "Произведено на Алтае"),

    ("5500647058", "СибКорм Каша Ассорти (мясо+рыба)", "Собаки", "Каши", "NAT", "Натуральность состава", "100% натуральный состав; без ГМО"),
    ("5500647058", "СибКорм Каша Ассорти (мясо+рыба)", "Собаки", "Каши", "ORIGIN", "Страна / регион производства", "Произведено на Алтае"),

    # --- Day 7 (2026-08-24): Сухой корм для собак ---
    ("5515836917", "Wanpy Dog беззерновой, курица", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain-Free; No added corn, wheat, soy or lentils"),
    ("5515836917", "Wanpy Dog беззерновой, курица", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "84% Animal Protein"),
    ("5515836917", "Wanpy Dog беззерновой, курица", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion"),

    ("5515838796", "Wanpy Dog беззерновой, лосось", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain-Free; No added corn, wheat, soy or lentils"),
    ("5515838796", "Wanpy Dog беззерновой, лосось", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "84% Animal Protein"),
    ("5515838796", "Wanpy Dog беззерновой, лосось", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion"),

    ("5524575556", "Wanpy Dog беззерновой, ягнёнок", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain-Free; No added corn, wheat, soy or lentils"),
    ("5524575556", "Wanpy Dog беззерновой, ягнёнок", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "84% Animal Protein"),
    ("5524575556", "Wanpy Dog беззерновой, ягнёнок", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion"),

    ("5529408722", "SIRIUS Premium индейка с овощами", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "41% животных ингредиентов"),
    ("5529408722", "SIRIUS Premium индейка с овощами", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Комплекс пробиотиков ProStor; крепкие кости; поддерживает иммунитет; здоровое сердце"),
    ("5529408722", "SIRIUS Premium индейка с овощами", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Premium Food"),

    ("5530070606", "Provipet с ягнёнком", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "Без искусственных ароматизаторов и усилителей вкуса"),
    ("5530070606", "Provipet с ягнёнком", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Высокая усвояемость; отличное пищеварение; крепкие зубы и кости; устойчивая иммунная система; здоровье кожи и шерсти"),
    ("5530070606", "Provipet с ягнёнком", "Собаки", "Сухой корм", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5555438805", "Provipet с индейкой", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "Без искусственных ароматизаторов и усилителей вкуса"),
    ("5555438805", "Provipet с индейкой", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Высокая усвояемость; отличное пищеварение; крепкие зубы и кости; устойчивая иммунная система; здоровье кожи и шерсти"),
    ("5555438805", "Provipet с индейкой", "Собаки", "Сухой корм", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5542850495", "AWARD Monoprotein лосось+лён+алоэ", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypo & Digestion"),
    ("5542850495", "AWARD Monoprotein лосось+лён+алоэ", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Monoprotein; Low Grain"),
    ("5542850495", "AWARD Monoprotein лосось+лён+алоэ", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% BIO"),
    ("5542850495", "AWARD Monoprotein лосось+лён+алоэ", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Fresh Meat"),
    ("5542850495", "AWARD Monoprotein лосось+лён+алоэ", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Superpremium Dog Food"),

    ("5556742789", "AlphaPet Superpremium Monoprotein белая рыба", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Монобелковая рецептура"),
    ("5556742789", "AlphaPet Superpremium Monoprotein белая рыба", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "30% свежей белой рыбы — ингредиент №1"),
    ("5556742789", "AlphaPet Superpremium Monoprotein белая рыба", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Чувствительное пищеварение; здоровая кожа и блестящая шерсть"),
    ("5556742789", "AlphaPet Superpremium Monoprotein белая рыба", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Superpremium"),

    ("5566163855", "Grandorf Fresh индейка с бататом", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic"),
    ("5566163855", "Grandorf Fresh индейка с бататом", "Собаки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
    ("5566163855", "Grandorf Fresh индейка с бататом", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Gut Health — Live Probiotics; Omega-3, Omega-6"),
    ("5566163855", "Grandorf Fresh индейка с бататом", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% Natural Dog Food"),
    ("5566163855", "Grandorf Fresh индейка с бататом", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5566163855", "Grandorf Fresh индейка с бататом", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Мясо первым ингредиентом"),

    ("5566490497", "AlphaPet Menu говядина", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Со свежим мясом"),
    ("5566490497", "AlphaPet Menu говядина", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% натуральные ингредиенты"),

    ("5568844559", "Vivere Maxi Adult ягнёнок", "Собаки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic Philosophy"),
    ("5568844559", "Vivere Maxi Adult ягнёнок", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% Natural Pet Food"),
    ("5568844559", "Vivere Maxi Adult ягнёнок", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Low Grain"),
    ("5568844559", "Vivere Maxi Adult ягнёнок", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "60% Animal Ingredients"),

    ("5574603661", "Purina Pro Plan Veterinary Diets EN Gastrointestinal", "Собаки", "Сухой корм", "VET", "Ветеринарная / лечебная диета", "Ветеринарная диета при расстройствах пищеварения"),

    ("5578347424", "Hilton Гипоаллергенный говядина", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Гипоаллергенный корм"),

    # --- Day 7 (2026-08-24): Влажный корм для собак ---
    ("5471075781", "Мнямс Максимум вкуса телятина", "Собаки", "Влажный корм", "NAT", "Натуральность состава", "Без консервантов; без ГМО"),

    ("5530369991", "Wanpy Dog консервы курица", "Собаки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Premium Quality"),
    ("5530369991", "Wanpy Dog консервы курица", "Собаки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Strengthens Bones; Healthy Digestion"),

    ("5530650366", "Wanpy Dog консервы говядина", "Собаки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Premium Quality"),
    ("5530650366", "Wanpy Dog консервы говядина", "Собаки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Strengthens Bones; Healthy Digestion"),

    ("5530650986", "Wanpy Dog консервы ягнёнок", "Собаки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Premium Quality"),
    ("5530650986", "Wanpy Dog консервы ягнёнок", "Собаки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Strengthens Bones; Healthy Digestion"),

    ("5544614194", "Royal Canin Mini Adult паштет", "Собаки", "Влажный корм", "ORIGIN", "Страна / регион производства", "Сделано в Австрии"),

    ("5549300374", "Pedigree паштет с телятиной (щенки)", "Собаки", "Влажный корм", "NAT", "Натуральность состава", "Натуральное мясо в составе"),

    # --- Day 7 (2026-08-24): Лакомства для собак ---
    ("5544395456", "Лёгкое говяжье для дрессировки", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральное"),
    ("5544395456", "Лёгкое говяжье для дрессировки", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Поддержка зубов и дёсен"),

    ("5555142584", "biohaus Медальоны из утки", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральный продукт; без ГМО; без глютена"),
    ("5555142584", "biohaus Медальоны из утки", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "99% владельцев рекомендуют"),

    ("5555380815", "Lucky Animals Luckers яблоко с говядиной", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральное лакомство"),

    ("5555663607", "biohaus Суши из кролика и трески", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральный продукт; без ГМО; без глютена"),
    ("5555663607", "biohaus Суши из кролика и трески", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "99% владельцев рекомендуют"),

    ("5556221688", "biohaus Суши из индейки и трески", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральный продукт; без ГМО; без глютена"),
    ("5556221688", "biohaus Суши из индейки и трески", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "99% владельцев рекомендуют"),

    ("5556538078", "biohaus Шашлычки куриные", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральный продукт; без ГМО; без глютена"),
    ("5556538078", "biohaus Шашлычки куриные", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "99% владельцев рекомендуют"),

    ("5556653006", "biohaus Уши кролика с уткой", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральный продукт; без ГМО; без глютена"),
    ("5556653006", "biohaus Уши кролика с уткой", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "99% владельцев рекомендуют"),

    ("5577356638", "TiTBiT Мясные сушки говядина", "Собаки", "Лакомства", "MEAT", "Происхождение и качество мяса", "50% мясопродуктов по сырому весу"),
    ("5577356638", "TiTBiT Мясные сушки говядина", "Собаки", "Лакомства", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5577356979", "TiTBiT Мясное печенье Бискотти", "Собаки", "Лакомства", "MEAT", "Происхождение и качество мяса", "50% мясопродуктов по сырому весу"),
    ("5577356979", "TiTBiT Мясное печенье Бискотти", "Собаки", "Лакомства", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5577357024", "TiTBiT (Biff) Dent сыр", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Для чистки зубов; массирует дёсны"),
    ("5577357024", "TiTBiT (Biff) Dent сыр", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Укрепление иммунитета"),
    ("5577357024", "TiTBiT (Biff) Dent сыр", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Для дрессировки"),

    ("5577357304", "TiTBiT (Biff) Dent утка", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Для чистки зубов; массирует дёсны"),
    ("5577357304", "TiTBiT (Biff) Dent утка", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Укрепление иммунитета"),
    ("5577357304", "TiTBiT (Biff) Dent утка", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Для дрессировки"),

    ("5577357359", "TiTBiT Соломка из лосося", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натурально"),
    ("5577357359", "TiTBiT Соломка из лосося", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Содержит Омега-3 и 6; полезно для шерсти и кожи"),
    ("5577357359", "TiTBiT Соломка из лосося", "Собаки", "Лакомства", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5577358728", "TiTBiT Кубики из лосося", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натурально"),
    ("5577358728", "TiTBiT Кубики из лосося", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Содержит Омега-3 и 6; полезно для шерсти и кожи"),
    ("5577358728", "TiTBiT Кубики из лосося", "Собаки", "Лакомства", "ORIGIN", "Страна / регион производства", "Сделано в России"),

    ("5577359119", "TiTBiT Крекер говяжий XL", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Способствует удалению зубного налёта и укреплению дёсен"),
    ("5577359119", "TiTBiT Крекер говяжий XL", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральный продукт, без искусственных красителей, консервантов и ароматизаторов"),

    ("5577402177", "Мистер Грызкин Говяжья трахея", "Собаки", "Лакомства", "NAT", "Натуральность состава", "Полностью натурально"),

    ("5585499483", "biohaus Сэндвич с уткой", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральный продукт; без ГМО; без глютена"),
    ("5585499483", "biohaus Сэндвич с уткой", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "99% владельцев рекомендуют"),

    # --- Day 7 (2026-08-24): Сухой корм для кошек ---
    ("5515840508", "Wanpy Cat беззерновой, курица", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain-Free; No added corn, wheat, soy or lentils"),
    ("5515840508", "Wanpy Cat беззерновой, курица", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "84% Animal Protein"),
    ("5515840508", "Wanpy Cat беззерновой, курица", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion"),

    ("5515844102", "Wanpy Cat беззерновой, лосось", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain-Free; No added corn, wheat, soy or lentils"),
    ("5515844102", "Wanpy Cat беззерновой, лосось", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "84% Animal Protein"),
    ("5515844102", "Wanpy Cat беззерновой, лосось", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion"),

    ("5542531112", "AWARD Monoprotein cat лосось+спаржа+ацерола", "Кошки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypo & Digestion"),
    ("5542531112", "AWARD Monoprotein cat лосось+спаржа+ацерола", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Monoprotein; Low Grain"),
    ("5542531112", "AWARD Monoprotein cat лосось+спаржа+ацерола", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% BIO"),
    ("5542531112", "AWARD Monoprotein cat лосось+спаржа+ацерола", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Fresh Meat"),
    ("5542531112", "AWARD Monoprotein cat лосось+спаржа+ацерола", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Superpremium Cat Food"),

    ("5542594928", "AJO Special Line Delicate Taste индейка", "Кошки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic View"),
    ("5542594928", "AJO Special Line Delicate Taste индейка", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Low Grain; Wheat/Corn Free"),

    ("5543934520", "AlphaPet Sensitive digestion", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "30% свежего мяса; 75% мясного белка"),
    ("5543934520", "AlphaPet Sensitive digestion", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Легкоусвояемый белок"),

    ("5543935300", "Goodwin индейка+утка", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Поддержка иммунной системы (розмарин, алоэ, зелёный чай)"),
    ("5543935300", "Goodwin индейка+утка", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "51%"),
    ("5543935300", "Goodwin индейка+утка", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),

    ("5543935895", "AlphaPet для крупных пород", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "30% свежего мяса; 75% мясного белка"),
    ("5543935895", "AlphaPet для крупных пород", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Легкоусвояемый белок"),

    ("5543936157", "Goodwin котята индейка", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Поддержка иммунной системы (розмарин, алоэ, зелёный чай)"),
    ("5543936157", "Goodwin котята индейка", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),

    ("5543936716", "AlphaPet стерилизованные цыплёнок", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "30% свежего мяса; 75% мясного белка"),
    ("5543936716", "AlphaPet стерилизованные цыплёнок", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Легкоусвояемый белок"),

    ("5554107233", "Primordial Cat Adult утка+индейка", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5554107233", "Primordial Cat Adult утка+индейка", "Кошки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic Nutrition"),
    ("5554107233", "Primordial Cat Adult утка+индейка", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% Natural"),
    ("5554107233", "Primordial Cat Adult утка+индейка", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "70% Animal Ingredients"),

    ("5554109198", "Primordial Neutered Exigent лосось+перепелка", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5554109198", "Primordial Neutered Exigent лосось+перепелка", "Кошки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic Nutrition"),
    ("5554109198", "Primordial Neutered Exigent лосось+перепелка", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% Natural"),
    ("5554109198", "Primordial Neutered Exigent лосось+перепелка", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Animal Ingredients"),

    ("5554109696", "Primordial Neutered Skin&Coat лосось+индейка", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5554109696", "Primordial Neutered Skin&Coat лосось+индейка", "Кошки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic Nutrition"),
    ("5554109696", "Primordial Neutered Skin&Coat лосось+индейка", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% Natural"),
    ("5554109696", "Primordial Neutered Skin&Coat лосось+индейка", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Animal Ingredients"),
    ("5554109696", "Primordial Neutered Skin&Coat лосось+индейка", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Здоровье кожи и шерсти (Skin&Coat)"),

    ("5556095810", "Monge Cat Monoprotein Sterilised утка", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Monoprotein"),
    ("5556095810", "Monge Cat Monoprotein Sterilised утка", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "First Ingredient №1 Duck"),
    ("5556095810", "Monge Cat Monoprotein Sterilised утка", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Enhanced palatability; sub-caloric content"),
    ("5556095810", "Monge Cat Monoprotein Sterilised утка", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Natural Superpremium"),

    ("5556862888", "Landor Sterilised индейка+утка", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Super Premium Food Class"),
    ("5556862888", "Landor Sterilised индейка+утка", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "Не содержит пшеницы; без ГМО"),

    ("5566319732", "AlphaPet Superpremium Monoprotein индейка (котята)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "30% свежего мяса; 75% мясного белка"),
    ("5566319732", "AlphaPet Superpremium Monoprotein индейка (котята)", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Легкоусвояемый белок"),

    ("5566555324", "AlphaPet Superpremium senior кролик+утка", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "30% свежего мяса; 75% мясного белка"),
    ("5566555324", "AlphaPet Superpremium senior кролик+утка", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Легкоусвояемый белок"),

    ("5566602205", "AlphaPet Superpremium говядина+печень", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "30% свежего мяса; 75% мясного белка"),
    ("5566602205", "AlphaPet Superpremium говядина+печень", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Легкоусвояемый белок"),

    ("5566632364", "AlphaPet Menu говядина (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Со свежим мясом"),
    ("5566632364", "AlphaPet Menu говядина (кошки)", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% натуральные ингредиенты"),

    ("5566657335", "AlphaPet Superpremium цыплёнок", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "30% свежего мяса; 75% мясного белка"),
    ("5566657335", "AlphaPet Superpremium цыплёнок", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Легкоусвояемый белок"),

    ("5568745202", "Karmy Hypoallergenic утка", "Кошки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Гипоаллергенный"),
    ("5568745202", "Karmy Hypoallergenic утка", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Улучшает работу пищеварительной системы; поддерживает здоровье кожи и шерсти"),
    ("5568745202", "Karmy Hypoallergenic утка", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Высокое содержание мяса"),

    ("5568879450", "Optima Nova беззерновой кролик+картофель", "Кошки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5568879450", "Optima Nova беззерновой кролик+картофель", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "75% свежего кролика + 30% дегидрированного + 5% гидролизата"),
    ("5568879450", "Optima Nova беззерновой кролик+картофель", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Omega 3&6; здоровье кожи и шерсти; высокая усвояемость"),

    ("5574288140", "Purina Pro Plan Veterinary Diets EN ST/OX Gastrointestinal (кошки)", "Кошки", "Сухой корм", "VET", "Ветеринарная / лечебная диета", "Для снижения проявлений кишечных расстройств"),

    ("5578339795", "Hilton Adult Sensitive кролик (кошки)", "Кошки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Гипоаллергенный корм"),
    ("5578339795", "Hilton Adult Sensitive кролик (кошки)", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Комплекс витаминов; сбалансированный рацион"),
    ("5578339795", "Hilton Adult Sensitive кролик (кошки)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Eco-Premium; Super Premium"),

    # --- Day 7 (2026-08-24): Влажный корм для кошек ---
    ("5484409205", "Gourmet Gold паштет курица", "Кошки", "Влажный корм", "ORIGIN", "Страна / регион производства", "Произведено во Франции"),
    ("5484413359", "Gourmet Gold паштет индейка", "Кошки", "Влажный корм", "ORIGIN", "Страна / регион производства", "Произведено во Франции"),
    ("5484418520", "Gourmet Gold паштет тунец", "Кошки", "Влажный корм", "ORIGIN", "Страна / регион производства", "Произведено во Франции"),

    ("5530649635", "Wanpy Cat тунец в собственном соку", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "No Preservatives, No Colorants"),
    ("5530649635", "Wanpy Cat тунец в собственном соку", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Real Meat"),
    ("5530649635", "Wanpy Cat тунец в собственном соку", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion, Easy to Absorb"),
    ("5530649635", "Wanpy Cat тунец в собственном соку", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),

    ("5530649765", "Wanpy Cat курица+креветки", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "No Preservatives, No Colorants"),
    ("5530649765", "Wanpy Cat курица+креветки", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Real Meat"),
    ("5530649765", "Wanpy Cat курица+креветки", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion, Easy to Absorb"),
    ("5530649765", "Wanpy Cat курица+креветки", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),

    ("5530650056", "Wanpy Cat курица+лосось", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "No Preservatives, No Colorants"),
    ("5530650056", "Wanpy Cat курица+лосось", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Real Meat"),
    ("5530650056", "Wanpy Cat курица+лосось", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion, Easy to Absorb"),
    ("5530650056", "Wanpy Cat курица+лосось", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),

    ("5530367150", "Wanpy Cat курица+гребешки", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "No Preservatives, No Colorants"),
    ("5530367150", "Wanpy Cat курица+гребешки", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Real Meat"),
    ("5530367150", "Wanpy Cat курица+гребешки", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion, Easy to Absorb"),
    ("5530367150", "Wanpy Cat курица+гребешки", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),

    ("5530650988", "Wanpy Cat рагу лосось+тунец (котята)", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "No Preservatives, No Colorants"),
    ("5530650988", "Wanpy Cat рагу лосось+тунец (котята)", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Real Meat"),
    ("5530650988", "Wanpy Cat рагу лосось+тунец (котята)", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion, Easy to Absorb"),
    ("5530650988", "Wanpy Cat рагу лосось+тунец (котята)", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),

    ("5530651154", "Wanpy Cat рагу курица+утка (котята)", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "No Preservatives, No Colorants"),
    ("5530651154", "Wanpy Cat рагу курица+утка (котята)", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Real Meat"),
    ("5530651154", "Wanpy Cat рагу курица+утка (котята)", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion, Easy to Absorb"),
    ("5530651154", "Wanpy Cat рагу курица+утка (котята)", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),

    ("5530651642", "Wanpy Cat курица+треска", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "No Preservatives, No Colorants"),
    ("5530651642", "Wanpy Cat курица+треска", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Real Meat"),
    ("5530651642", "Wanpy Cat курица+треска", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion, Easy to Absorb"),
    ("5530651642", "Wanpy Cat курица+треска", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),

    ("5530653845", "Wanpy Cat курица+тунец", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "No Preservatives, No Colorants"),
    ("5530653845", "Wanpy Cat курица+тунец", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Real Meat"),
    ("5530653845", "Wanpy Cat курица+тунец", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion, Easy to Absorb"),
    ("5530653845", "Wanpy Cat курица+тунец", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),

    ("5530654016", "Wanpy Cat курица+краб", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "No Preservatives, No Colorants"),
    ("5530654016", "Wanpy Cat курица+краб", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Real Meat"),
    ("5530654016", "Wanpy Cat курица+краб", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Healthy Digestion, Easy to Absorb"),
    ("5530654016", "Wanpy Cat курица+краб", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Super Premium"),

    ("5551551944", "Royal Canin Kitten паштет", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "No Preservatives"),
    ("5551551944", "Royal Canin Kitten паштет", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Health Nutrition since 1968"),

    ("5566574938", "Schesir After Dark курица+перепелиное яйцо", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5566598587", "Schesir After Dark курица", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5566620215", "Schesir After Dark курица+говядина", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5566642607", "Schesir After Dark курица+утка", "Кошки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),

    # --- Day 7 (2026-08-24): Лакомства для кошек ---
    ("5542724663", "АШАН Pets Хрустящие сердечки индейка+сливки", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Без добавления искусственных ароматизаторов"),
    ("5542724663", "АШАН Pets Хрустящие сердечки индейка+сливки", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Витамины и минералы в составе"),

    ("5543308856", "Ферма кота Федора палочки кролик", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Без ГМО; без сои"),
    ("5543534231", "Ферма кота Федора палочки лосось", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Без ГМО; без сои"),
    ("5543551335", "Ферма кота Федора палочки утка", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Без ГМО; без сои"),

    ("5549365822", "Dreamies Молочный коктейль", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Без искусственных ароматизаторов"),
    ("5549365822", "Dreamies Молочный коктейль", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Витамины и минералы; высокое содержание белка"),

    ("5579052755", "AWARD Special Care утка+тунец", "Кошки", "Лакомства", "NAT", "Натуральность состава", "100% BIO"),
    ("5579052755", "AWARD Special Care утка+тунец", "Кошки", "Лакомства", "MEAT", "Происхождение и качество мяса", "Fresh Meat"),
    ("5579052755", "AWARD Special Care утка+тунец", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Чувствительное пищеварение (Nutraceutical Formula)"),
    ("5579052755", "AWARD Special Care утка+тунец", "Кошки", "Лакомства", "PREMIUM", "Премиальное позиционирование", "Quality for Champions"),

    ("5579058668", "AWARD Special Care курица", "Кошки", "Лакомства", "NAT", "Натуральность состава", "100% BIO"),
    ("5579058668", "AWARD Special Care курица", "Кошки", "Лакомства", "MEAT", "Происхождение и качество мяса", "Fresh Meat"),
    ("5579058668", "AWARD Special Care курица", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Профилактика мочекаменной болезни (Urinary)"),
    ("5579058668", "AWARD Special Care курица", "Кошки", "Лакомства", "PREMIUM", "Премиальное позиционирование", "Quality for Champions"),

    ("5579066625", "AWARD Special Care лосось+тунец", "Кошки", "Лакомства", "NAT", "Натуральность состава", "100% BIO"),
    ("5579066625", "AWARD Special Care лосось+тунец", "Кошки", "Лакомства", "MEAT", "Происхождение и качество мяса", "Fresh Meat"),
    ("5579066625", "AWARD Special Care лосось+тунец", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Здоровая кожа и шерсть"),
    ("5579066625", "AWARD Special Care лосось+тунец", "Кошки", "Лакомства", "PREMIUM", "Премиальное позиционирование", "Quality for Champions"),

    ("5582640835", "Finesse Licko курица+тыква", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Без искусственных красителей"),
    ("5582640835", "Finesse Licko курица+тыква", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "For digestive care (тыквенная клетчатка)"),

    ("5584169750", "Murkel курица+рыбий жир+козье молоко", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Натуральный состав"),
    ("5584169750", "Murkel курица+рыбий жир+козье молоко", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Для поддержания пищеварения/иммунитета/мягкой шёрстки"),

    ("5584218651", "Murkel льняное семя+курица+печень", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Натуральный состав"),
    ("5584218651", "Murkel льняное семя+курица+печень", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Для иммунитета/зрения и обмена веществ"),

    ("5584264918", "Murkel тунец+молочные продукты", "Кошки", "Лакомства", "NAT", "Натуральность состава", "Натуральный состав"),
    ("5584264918", "Murkel тунец+молочные продукты", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Для поддержания пищеварения/иммунитета"),

    # --- Day 8 (2026-09-02): Сухой корм для кошек ---
    ("5658764220", "Родные корма курица+ягоды (кошки)", "Кошки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Мясо ингредиент №1; 63% белка животного происхождения"),
    ("5658764220", "Родные корма курица+ягоды (кошки)", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Премиум"),
    ("5658764220", "Родные корма курица+ягоды (кошки)", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Здоровая кожа и шерсть; контроль веса; комплекс пробиотиков"),

    ("5670474832", "AWARD Hairball & Indoor утка+индейка", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% BIO"),
    ("5670474832", "AWARD Hairball & Indoor утка+индейка", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Hairball Control; Nutraceutical Formula"),
    ("5670474832", "AWARD Hairball & Indoor утка+индейка", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Superpremium Cat Food; Quality for Champions"),

    ("5670477242", "AWARD Urinary курица", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% BIO"),
    ("5670477242", "AWARD Urinary курица", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Urinary Health (профилактика МКБ); Nutraceutical Formula"),
    ("5670477242", "AWARD Urinary курица", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Superpremium Cat Food; Quality for Champions"),

    ("5670498757", "AWARD Indoor Big cats говядина+курица", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% BIO"),
    ("5670498757", "AWARD Indoor Big cats говядина+курица", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Nutraceutical Formula"),
    ("5670498757", "AWARD Indoor Big cats говядина+курица", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Superpremium Cat Food; Quality for Champions"),

    ("5670499419", "AWARD Sterilized индейка+курица", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% BIO"),
    ("5670499419", "AWARD Sterilized индейка+курица", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Urinary Health; Nutraceutical Formula"),
    ("5670499419", "AWARD Sterilized индейка+курица", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Superpremium Cat Food; Quality for Champions"),

    ("5670501608", "AWARD Sterilized белая рыба", "Кошки", "Сухой корм", "NAT", "Натуральность состава", "100% BIO"),
    ("5670501608", "AWARD Sterilized белая рыба", "Кошки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Urinary Health; Nutraceutical Formula"),
    ("5670501608", "AWARD Sterilized белая рыба", "Кошки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Superpremium Cat Food; Quality for Champions"),

    # --- Day 8 (2026-09-02): Влажный корм для собак ---
    ("5669273502", "Yummy Мясной Деликатес индейка", "Собаки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Без зерна; монобелковый"),
    ("5670094001", "Мнямс Максимум вкуса говядина+морковь", "Собаки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Не содержит злаков"),
    ("5670712472", "Мнямс Максимум вкуса телятина", "Собаки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Не содержит злаков"),
    ("5670714813", "Мнямс Сильный иммунитет утка+перепел", "Собаки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Не содержит злаков"),
    ("5670714813", "Мнямс Сильный иммунитет утка+перепел", "Собаки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Сильный иммунитет"),
    ("5670718836", "Мнямс мусс курица+печень", "Собаки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Не содержит злаков"),
    ("5670720625", "Мнямс Террин по-версальски телятина+ветчина", "Собаки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Не содержит злаков"),
    ("5671464738", "Мнямс мусс утка", "Собаки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Не содержит злаков"),
    ("5671505639", "Мнямс Мусс по-французски индейка+цукини", "Собаки", "Влажный корм", "GRAIN", "Беззерновой / монобелковый состав", "Не содержит злаков"),

    # --- Day 8 (2026-09-02): Сухой корм для собак ---
    ("5623144132", "Farmina N&D Prime ягнёнок+черника (собаки)", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5623144132", "Farmina N&D Prime ягнёнок+черника (собаки)", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "Nutrition System for Carnivores"),
    ("5623144132", "Farmina N&D Prime ягнёнок+черника (собаки)", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Italian formula"),

    ("5634101309", "GRANDORF Veterinary Diet Hypoallergenic кролик", "Собаки", "Сухой корм", "VET", "Ветеринарная / лечебная диета", "Approved by Vets; ветеринарная диета"),
    ("5634101309", "GRANDORF Veterinary Diet Hypoallergenic кролик", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "При пищевой аллергии и непереносимости"),
    ("5634101309", "GRANDORF Veterinary Diet Hypoallergenic кролик", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Gluten & Grain Free"),
    ("5634101309", "GRANDORF Veterinary Diet Hypoallergenic кролик", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Антистресс; пробиотики/постбиотики для ЖКТ и иммунитета; здоровье кожи и шерсти"),

    ("5662623115", "GRANDORF Fresh лосось+батат", "Собаки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
    ("5662623115", "GRANDORF Fresh лосось+батат", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5662623115", "GRANDORF Fresh лосось+батат", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Живые пробиотики"),

    ("5662623376", "GRANDORF Fresh индейка+батат", "Собаки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
    ("5662623376", "GRANDORF Fresh индейка+батат", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5662623376", "GRANDORF Fresh индейка+батат", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Живые пробиотики"),

    ("5662626493", "GRANDORF Fresh утка+батат", "Собаки", "Сухой корм", "HOLISTIC", "Холистик-позиционирование", "Holistic"),
    ("5662626493", "GRANDORF Fresh утка+батат", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "Grain Free"),
    ("5662626493", "GRANDORF Fresh утка+батат", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Живые пробиотики"),

    ("5670255238", "Royal Canin Expert Neutered Adult мини (собаки)", "Собаки", "Сухой корм", "VET", "Ветеринарная / лечебная диета", "№1 по рекомендации ветеринарных врачей; ветеринарный рацион Expert"),
    ("5670255238", "Royal Canin Expert Neutered Adult мини (собаки)", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Royal Canin"),

    ("5670469697", "AWARD HYPO белая рыба (собаки)", "Собаки", "Сухой корм", "HYPO", "Гипоаллергенность / чувствит. пищеварение", "Hypoallergenic"),
    ("5670469697", "AWARD HYPO белая рыба (собаки)", "Собаки", "Сухой корм", "NAT", "Натуральность состава", "100% BIO"),
    ("5670469697", "AWARD HYPO белая рыба (собаки)", "Собаки", "Сухой корм", "MEAT", "Происхождение и качество мяса", "65% животного белка"),
    ("5670469697", "AWARD HYPO белая рыба (собаки)", "Собаки", "Сухой корм", "GRAIN", "Беззерновой / монобелковый состав", "0% глютен; 0% соя"),
    ("5670469697", "AWARD HYPO белая рыба (собаки)", "Собаки", "Сухой корм", "FUNC", "Функциональная польза для здоровья", "Long-term Energy Boost; Nutraceutical Formula"),
    ("5670469697", "AWARD HYPO белая рыба (собаки)", "Собаки", "Сухой корм", "PREMIUM", "Премиальное позиционирование", "Superpremium Dog Food; Quality for Champions"),

    # --- Day 8 (2026-09-02): Влажный корм для кошек ---
    ("5654017462", "AlphaPet WOW кролик+потрошки", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Чувствительное пищеварение; здоровое пищеварение"),
    ("5654017462", "AlphaPet WOW кролик+потрошки", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "Fresh Meat (Rich Meat Profile)"),
    ("5654017462", "AlphaPet WOW кролик+потрошки", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Superpremium"),
    ("5654017462", "AlphaPet WOW кролик+потрошки", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Комплекс AlphaPetBIO"),

    ("5664781811", "P.E.P.P.O. курица", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Натуральные ингредиенты"),
    ("5664781811", "P.E.P.P.O. курица", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Поддержка иммунитета"),
    ("5664781811", "P.E.P.P.O. курица", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "85% мяса в каждом кусочке"),

    ("5664784502", "P.E.P.P.O. ягнёнок (котята)", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Натуральные ингредиенты"),
    ("5664784502", "P.E.P.P.O. ягнёнок (котята)", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Поддержка иммунитета"),
    ("5664784502", "P.E.P.P.O. ягнёнок (котята)", "Кошки", "Влажный корм", "MEAT", "Происхождение и качество мяса", "85% мяса в каждом кусочке"),

    ("5668947557", "Вкусная Миска говядина", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Без красителей; без ГМО"),
    ("5668947557", "Вкусная Миска говядина", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "С витаминами и минералами"),

    ("5669095176", "Вкусная Миска кролик", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Без красителей; без ГМО"),
    ("5669095176", "Вкусная Миска кролик", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "С витаминами и минералами"),

    ("5670228940", "Whiskas Особое удовольствие индейка+печень+овощи", "Кошки", "Влажный корм", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Особое удовольствие"),
    ("5670237889", "SHEBA кролик+индейка", "Кошки", "Влажный корм", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Ломтики в соусе"),
    ("5670251820", "SHEBA курица", "Кошки", "Влажный корм", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Ломтики в соусе"),

    ("5670328613", "Perfect Fit индейка", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Чувствительное пищеварение; здоровая микрофлора кишечника"),
    ("5670328613", "Perfect Fit индейка", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Без добавления пшеницы и сои"),
    ("5670328613", "Perfect Fit индейка", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Разработано со специалистами Waltham Petcare Science Institute"),

    ("5670392893", "Royal Canin Sterilised соус (кошки)", "Кошки", "Влажный корм", "PREMIUM", "Премиальное позиционирование", "Health Nutrition since 1968"),

    ("5670714887", "Мнямс Ферма кота Фёдора лосось+креветки+форель", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Без ГМО; без сои"),
    ("5670716904", "Мнямс Ферма кота Фёдора индейка (котята)", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Без ГМО; без сои"),
    ("5670719609", "Мнямс Ферма кота Фёдора телятина (котята)", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Без ГМО; без сои"),
    ("5670721156", "Мнямс Ферма кота Фёдора говядина", "Кошки", "Влажный корм", "NAT", "Натуральность состава", "Без ГМО; без сои"),

    ("5671557401", "Мнямс кусочки в соусе утка (кошки)", "Кошки", "Влажный корм", "FUNC", "Функциональная польза для здоровья", "Профилактика мочекаменной болезни; поддержание оптимальной кондиции"),

    # --- Day 8 (2026-09-02): Лакомства для собак ---
    ("5632614860", "INABA Neji говядина+курица+сыр", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "На гриле, со вкусом сыра"),
    ("5632615231", "INABA Churu Bee курица (суставы)", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Поддержание здоровья суставов"),
    ("5632624539", "INABA Churu курица (кожа и шерсть, собаки)", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Здоровье кожи и шерсти"),
    ("5654589093", "INABA Churu Ball курица+овощи", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Кремовая начинка"),
    ("5654591896", "INABA Churu курица (ЖКТ, собаки)", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Здоровье ЖКТ"),
    ("5654592771", "INABA Churu Ball курица", "Собаки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Кремовая начинка"),
    ("5654599026", "INABA Tobitsuku курица (кожа, собаки)", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Поддержание здоровья кожи"),
    ("5654599608", "INABA Tobitsuku курица (ЖКТ, собаки)", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Поддержание здоровья ЖКТ"),

    ("5670711483", "Мнямс DENTAL лосось", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Профилактика образования зубного камня; массаж дёсен"),
    ("5670715414", "Мнямс Колбаски аппетитные", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральное лакомство; без добавок и красителей"),
    ("5670715414", "Мнямс Колбаски аппетитные", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Улучшает пищеварение"),
    ("5670717416", "Мнямс Грудки куриные", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральное лакомство; без добавок и красителей"),
    ("5670717416", "Мнямс Грудки куриные", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Улучшает пищеварение"),
    ("5670718664", "Мнямс Филе миньон по-фламандски", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральное лакомство; без добавок и красителей"),
    ("5670718664", "Мнямс Филе миньон по-фламандски", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Улучшает пищеварение"),
    ("5670720986", "Мнямс Рубец говяжий", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральное лакомство; без добавок и красителей"),
    ("5670720986", "Мнямс Рубец говяжий", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Улучшает пищеварение"),
    ("5671459555", "Мнямс DENTAL говядина (крупные породы)", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Профилактика образования зубного камня; массаж дёсен"),
    ("5671490575", "Мнямс DENTAL зубные косточки", "Собаки", "Лакомства", "DENT", "Стоматологическая польза", "Профилактика образования зубного камня; массаж дёсен"),
    ("5671509904", "Мнямс Трахея говяжья", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральное лакомство; без добавок и красителей"),
    ("5671509904", "Мнямс Трахея говяжья", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Улучшает пищеварение"),
    ("5671561038", "Мнямс Колбаски из баранины", "Собаки", "Лакомства", "NAT", "Натуральность состава", "100% натуральное лакомство; без добавок и красителей"),
    ("5671561038", "Мнямс Колбаски из баранины", "Собаки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Улучшает пищеварение"),

    # --- Day 8 (2026-09-02): Лакомства для кошек ---
    ("5632614370", "INABA CIAO Paste курица+лактобактерии", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Здоровье кожи и шерсти"),
    ("5632614943", "INABA CIAO Churu тунец Кацуо", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Лакомство-пюре"),
    ("5632615075", "INABA Yaki Sasami курица+краб", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Со вкусом краба"),
    ("5632615106", "INABA Yaki Miksu ассорти", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Со вкусом кацуобуси, куриного бульона и кальмара"),
    ("5632619730", "INABA CIAO Churu курица (кожа и шерсть, кошки)", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Здоровье кожи и шерсти"),
    ("5654586643", "INABA CIAO Churu тунец Магуро", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Лакомство-пюре"),
    ("5654587557", "INABA Tobitsuku куриный бульон (кошки)", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Со вкусом куриного бульона"),
    ("5654590710", "INABA CIAO Churu курица (pH мочи)", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Поддержание pH мочи"),
    ("5654591217", "INABA CIAO Pon Churu тунец Магуро суфле", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Суфле-паштет"),
    ("5654591653", "INABA CIAO Churu курица+краб", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Лакомство-пюре"),
    ("5654592626", "INABA CIAO Churu тунец+гребешок", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Лакомство-пюре"),
    ("5654594462", "INABA CIAO Churu тунец+лобстер", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Лакомство-пюре"),
    ("5654596008", "INABA CIAO Churu курица (мочеполовая система)", "Кошки", "Лакомства", "FUNC", "Функциональная польза для здоровья", "Здоровье мочеполовой системы"),
    ("5654598009", "INABA CIAO Churu курица+кальмар", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Лакомство-пюре"),
    ("5654598442", "INABA Tobitsuku морской гребешок", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Со вкусом морского гребешка"),
    ("5654598890", "INABA CIAO Churu курица (котята)", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Лакомство-пюре"),
    ("5654599052", "INABA Yaki Sasami курица+говядина", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Со вкусом говядины"),
    ("5654599091", "INABA CIAO Pon Churu курица+гребешок суфле", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Суфле-паштет"),
    ("5654599124", "INABA CIAO Churu курица+креветка", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Лакомство-пюре"),
    ("5670711038", "Мнямс крем утка+тунец Кацуо", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Нежное пюре-лакомство"),
    ("5671503941", "Мнямс крем курица", "Кошки", "Лакомства", "TASTE", "Вкусовая привлекательность / поведенческая польза", "Нежное пюре-лакомство"),
]

for i, row in enumerate(claim_rows, start=2):
    for j, val in enumerate(row, start=1):
        c = ws2.cell(row=i, column=j, value=val)
        c.font = base_font
        c.alignment = wrap
        c.border = border

autosize(ws2, [12, 30, 12, 12, 10, 32, 55])
ws2.freeze_panes = "A2"

# Пересчитываем "Кол-во категорий клеймов" на листе "Примеры" напрямую из
# claim_rows (а не из вручную поддерживаемого cats_per_product) — так это
# поле физически не может разойтись с реальными клеймами ниже по мере роста
# выборки и правок задним числом.
_ncat_by_oid = {}
for _row in claim_rows:
    _ncat_by_oid.setdefault(_row[0], set()).add(_row[4])
for _i, _p in enumerate(products):
    ws.cell(row=hrow + 1 + _i, column=13, value=len(_ncat_by_oid.get(_p["oid"], set())))

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
# завода — не более одного). "—" в тегах не бывает: там, где на лицевой стороне упаковки
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
    "5327539554": [], "5327528788": [],  # не указан на лицевой стороне упаковки
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

    # ---- Day 5 (2026-08-04) ----
    "5345830359": ["Индейка"], "5358300410": [],
    "5357288481": ["Перепел"],
    "5358242931": ["Говядина"], "5358240926": ["Баранина"],
    "5359892602": ["Индейка", "Лосось"], "5359453152": ["Говядина", "Лосось"], "5359660535": [],
    "5359822345": ["Ягнёнок", "Индейка"],
    "5357563517": ["Курица"], "5357853793": ["Говядина"],
    "5331257302": ["Индейка"],
    "5358325770": ["Курица", "Треска", "Лосось", "Кролик", "Тунец", "Утка", "Гребешок"],
    "5358242441": ["Тунец"], "5358240663": ["Говядина"], "5358241683": ["Курица"],
    "5358240549": [],

    # ---- Day 6 (2026-08-12) ----
    "5430117217": ["Говядина"], "5423933879": ["Индейка"], "5423741999": ["Свинина", "Говядина"],
    "5432594014": ["Рыба"], "5419467431": [], "5418435027": [], "5417460160": ["Говядина"],
    "5432515655": ["Курица", "Печень"], "5431773123": ["Кролик", "Лосось"],
    "5384510562": ["Утка", "Лосось"], "5429960930": ["Курица", "Индейка", "Утка"],
    "5430728946": [], "5430727673": [], "5430194418": ["Курица", "Треска"],
    "5424780704": ["Лосось", "Треска"], "5424781094": ["Ягнёнок"],
    "5424780086": ["Курица"], "5424781029": ["Курица"], "5424783286": ["Ягнёнок"],
    "5424776640": ["Перепел"], "5424785573": ["Кабан"], "5424783287": ["Ягнёнок"],
    "5424784781": ["Утка"], "5424782869": ["Курица"],
    "5428395063": ["Говядина"], "5428925488": ["Говядина"],
    "5429259395": ["Творог"], "5431102190": ["Телятина"],
    "5430211436": ["Телятина"], "5432604231": ["Индейка"], "5432823007": ["Индейка"],

    # ---- Доразбор 12.08 (найдены аудитом ?sorting=new) ----
    "5433103013": [], "5433102619": ["Кролик"],
    "5433103250": ["Перепел", "Курица"], "5433103558": ["Кролик", "Говядина"],
    "5433104303": ["Индейка"], "5433102991": ["Кролик"],
    "5433102865": ["Калтык и язык"], "5433103171": ["Говядина", "Ягнёнок"],
    "5436654518": ["Утка"], "5424786354": ["Курица"], "5424780417": ["Ягнёнок"],

    # ---- Day 7 (2026-08-24): Каша для собак ----
    "5542503462": ["Курица"], "5445105149": ["Рыба"], "5445127064": [],
    "5456129061": [], "5456129661": [], "5456131380": [], "5456131513": [],
    "5479813617": [], "5500647058": [],

    # ---- Day 7 (2026-08-24): Сухой корм для собак ----
    "5515836917": ["Курица"], "5515838796": ["Лосось"], "5524575556": ["Ягнёнок"],
    "5529408722": ["Индейка"], "5530070606": ["Ягнёнок"], "5555438805": ["Индейка"],
    "5542850495": ["Лосось"], "5556742789": ["Белая рыба"], "5566163855": ["Индейка"],
    "5566490497": ["Говядина"], "5568844559": ["Ягнёнок"], "5574603661": [],
    "5578347424": ["Говядина"],

    # ---- Day 7 (2026-08-24): Влажный корм для собак ----
    "5471075781": ["Телятина"], "5530369991": ["Курица"], "5530650366": ["Говядина"],
    "5530650986": ["Ягнёнок"], "5544614194": [], "5549300374": ["Телятина"],
    "5556733970": ["Индейка"], "5556739255": ["Телятина", "Рубец"],

    # ---- Day 7 (2026-08-24): Лакомства для собак ----
    "5544395456": ["Говядина"], "5555142584": ["Утка"], "5555380815": ["Говядина"],
    "5555663607": ["Кролик", "Треска"], "5556221688": ["Индейка", "Треска"],
    "5556538078": ["Курица"], "5556653006": ["Кролик", "Утка"],
    "5577356638": ["Говядина"], "5577356979": [], "5577357024": ["Сыр"],
    "5577357304": ["Утка"], "5577357359": ["Лосось"], "5577358728": ["Лосось"],
    "5577359119": ["Говядина"], "5577402177": ["Говядина"], "5585499483": ["Утка"],

    # ---- Day 7 (2026-08-24): Сухой корм для кошек ----
    "5515840508": ["Курица"], "5515844102": ["Лосось"], "5542531112": ["Лосось"],
    "5542594928": ["Индейка"], "5543934520": [], "5543935300": ["Индейка", "Утка"],
    "5543935895": [], "5543936157": ["Индейка"], "5543936716": ["Курица"],
    "5554107233": ["Утка", "Индейка"], "5554109198": ["Лосось", "Перепел"],
    "5554109696": ["Лосось", "Индейка"], "5555801791": ["Говядина"],
    "5556095810": ["Утка"], "5556862888": ["Индейка", "Утка"],
    "5566319732": ["Индейка"], "5566555324": ["Кролик", "Утка"],
    "5566602205": ["Говядина", "Печень"], "5566632364": ["Говядина"],
    "5566657335": ["Курица"], "5568745202": ["Утка"], "5568879450": ["Кролик"],
    "5574288140": [], "5578339795": ["Кролик"],

    # ---- Day 7 (2026-08-24): Влажный корм для кошек ----
    "5484409205": ["Курица"], "5484413359": ["Индейка"], "5484418520": ["Тунец"],
    "5530649635": ["Тунец"], "5530649765": ["Курица", "Креветки"],
    "5530650056": ["Курица", "Лосось"], "5530367150": ["Курица", "Гребешок"],
    "5530650988": ["Лосось", "Тунец"], "5530651154": ["Курица", "Утка"],
    "5530651642": ["Курица", "Треска"], "5530653845": ["Курица", "Тунец"],
    "5530654016": ["Курица", "Краб"], "5538615147": ["Индейка", "Курица"],
    "5543025974": ["Курица", "Говядина", "Кролик", "Утка"], "5544350676": ["Говядина"],
    "5544357753": [], "5551551944": [], "5566574938": ["Курица", "Перепелиное яйцо"],
    "5566598587": ["Курица"], "5566620215": ["Курица", "Говядина"],
    "5566642607": ["Курица", "Утка"],

    # ---- Day 7 (2026-08-24): Лакомства для кошек ----
    "5542724663": ["Индейка"], "5543308856": ["Кролик"], "5543534231": ["Лосось"],
    "5543551335": ["Утка"], "5549365822": [], "5579052755": ["Утка", "Тунец"],
    "5579058668": ["Курица"], "5579066625": ["Лосось", "Тунец"], "5582640835": ["Курица"],
    "5584169750": ["Курица"], "5584218651": ["Курица", "Печень"], "5584264918": ["Тунец"],

    # ---- Day 8 (2026-09-02) ----
    "5658764220": ["Курица"],
    "5670474832": ["Утка", "Индейка"], "5670477242": ["Курица"], "5670498757": ["Говядина", "Курица"],
    "5670499419": ["Индейка", "Курица"], "5670501608": ["Белая рыба"],
    "5669273502": ["Индейка"], "5670094001": ["Говядина"], "5670712472": ["Телятина"],
    "5670714813": ["Утка", "Перепел"], "5670718836": ["Курица", "Печень"], "5670720625": ["Телятина"],
    "5671464738": ["Утка"], "5671505639": ["Индейка"],
    "5623144132": ["Ягнёнок"], "5634101309": ["Кролик"], "5662623115": ["Лосось"],
    "5662623376": ["Индейка"], "5662626493": ["Утка"], "5670255238": [], "5670469697": ["Белая рыба"],
    "5654017462": ["Кролик"], "5664781811": ["Курица"], "5664784502": ["Ягнёнок"],
    "5668947557": ["Говядина"], "5669095176": ["Кролик"], "5670228940": ["Индейка", "Печень"],
    "5670237889": ["Кролик", "Индейка"], "5670251820": ["Курица"], "5670328613": ["Индейка"],
    "5670392893": [], "5670714887": ["Лосось", "Креветки", "Форель"], "5670716904": ["Индейка"],
    "5670719609": ["Телятина"], "5670721156": ["Говядина"], "5671557401": ["Утка"],
    "5632614860": ["Говядина", "Курица"], "5632615231": ["Курица"], "5632624539": ["Курица"],
    "5654589093": ["Курица"], "5654591896": ["Курица"], "5654592771": ["Курица"],
    "5654599026": ["Курица"], "5654599608": ["Курица"],
    "5670711483": ["Лосось"], "5670715414": [], "5670717416": ["Курица"], "5670718664": [],
    "5670720986": ["Говядина"], "5671459555": ["Говядина"], "5671490575": [], "5671509904": ["Говядина"],
    "5671561038": ["Баранина"],
    "5632614370": ["Курица"], "5632614943": ["Тунец"], "5632615075": ["Курица"],
    "5632615106": ["Курица", "Кальмар"], "5632619730": ["Курица"], "5654586643": ["Тунец"],
    "5654587557": ["Курица"], "5654590710": ["Курица"], "5654591217": ["Тунец"],
    "5654591653": ["Курица", "Краб"], "5654592626": ["Тунец", "Гребешок"], "5654594462": ["Тунец", "Лобстер"],
    "5654596008": ["Курица"], "5654598009": ["Курица", "Кальмар"], "5654598442": ["Гребешок"],
    "5654598890": ["Курица"], "5654599052": ["Курица", "Говядина"], "5654599091": ["Курица", "Гребешок"],
    "5654599124": ["Курица", "Креветки"], "5670711038": ["Утка", "Тунец"], "5671503941": ["Курица"],
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
    "5358300410", "5359660535", "5358240549",
    "5419467431", "5418435027", "5430728946", "5430727673",
    "5433103013",
    "5445127064", "5456129061", "5456129661", "5456131380", "5456131513",
    "5479813617", "5500647058",
    "5574603661",
    "5544614194",
    "5577356979",
    "5543934520", "5543935895", "5574288140",
    "5544357753", "5551551944",
    "5549365822",
    "5670255238", "5670392893", "5670715414", "5670718664", "5671490575",
}
PROTEIN_TOTAL_SKUS = len(PROTEIN_TAGS) + len(EXTRA_PROTEIN_ONLY)

INGREDIENT_TAGS = {
    "5277005073": ["Батат", "Омега-3/6", "Глюкозамин/хондроитин", "Пробиотики"],
    "5276044539": ["Яблоко", "Артишок"],
    "4098300679": ["Трава (овёс)", "Витамины", "Фолиевая кислота"],
    "5294272736": ["Морская капуста (водоросли)", "Морковь"],
    "5306421483": ["Рис"], "5306617135": ["Рис"], "5327569955": ["Картофель"],
    "5306618148": ["Рис"], "5259705799": ["Овощи"],
    "5294353414": ["Рис"], "5294208319": ["Морская капуста (водоросли)"],
    "5294159811": ["Морская капуста (водоросли)"], "5294111572": ["Морская капуста (водоросли)"],
    "5306470815": ["Рис"], "5306424879": ["Рис"], "5328154887": ["Омега-3/6"],

    # ---- Day 5 (2026-08-04) ----
    "5345830359": ["Пребиотики"],
    "5358300410": ["Таурин", "Кальций/Фосфор", "Витамин D", "Жирные кислоты"],
    "5357288481": ["Тыква"],
    "5359822345": ["Омега-3/6"],
    "5357563517": ["Морковь", "Овёс"], "5357853793": ["Морковь", "Овёс"],
    "5331257302": ["Клюква", "Пробиотики"],
    "5358325770": ["Катнип", "Рыбий жир", "Пребиотики", "Криль"],
    "5358242441": ["Таурин"], "5358240663": ["Таурин"], "5358241683": ["Таурин"],
    "5358240549": ["Овёс", "Рожь", "Пшеница", "Тимофеевка", "Витамины"],

    # ---- Day 6 (2026-08-12) ----
    "5424780704": ["Дыня канталупа"],
    "5424781094": ["Киноа", "Брокколи", "Спаржа"],
    "5424780086": ["Полба", "Овёс", "Гранат"],
    "5424781029": ["Полба", "Овёс", "Гранат"],
    "5424783286": ["Полба", "Овёс", "Черника"],
    "5424776640": ["Тыква", "Гранат"],
    "5424785573": ["Тыква", "Яблоко"],
    "5424783287": ["Тыква", "Черника"],
    "5424784781": ["Киноа", "Брокколи", "Спаржа"],
    "5424782869": ["Гранат"],
    "5428395063": ["Мята", "Хлорофилл", "Ca/P"],

    # ---- Доразбор 12.08 ----
    "5433102619": ["Клюква", "Брокколи", "Ромашка"],
    "5433103250": ["Овощи"], "5433103558": ["Болгарский перец"],
    "5433104303": ["Яблоки", "Отруби"],
    "5424786354": ["Гранат"], "5424780417": ["Полба", "Овёс", "Черника"],

    # ---- Day 7 (2026-08-24): Каша для собак ----
    "5445127064": ["Гречка"],
    "5456129061": ["Злаки"], "5456129661": ["Рис"],
    "5456131380": ["Злаки"], "5456131513": ["Гречка"],

    # ---- Day 7 (2026-08-24): Сухой корм для собак ----
    "5515836917": ["Черника", "Горошек"],
    "5529408722": ["Морковь", "Сельдерей", "Пробиотики"],
    "5542850495": ["Семена льна", "Алоэ вера"],
    "5566163855": ["Батат"],
    "5568844559": ["Рис", "Горох"],

    # ---- Day 7 (2026-08-24): Лакомства для собак ----
    "5555380815": ["Яблоко"],
    "5577357359": ["Омега-3/6"], "5577358728": ["Омега-3/6"],

    # ---- Day 7 (2026-08-24): Сухой корм для кошек ----
    "5542531112": ["Спаржа", "Ацерола"],
    "5568879450": ["Картофель", "Омега-3/6"],

    # ---- Day 7 (2026-08-24): Влажный корм для кошек ----
    "5543025974": ["Морковь", "Кошачья мята", "Клюква"],

    # ---- Day 7 (2026-08-24): Лакомства для кошек ----
    "5542724663": ["Сливки"], "5582640835": ["Тыквенная клетчатка"],
    "5584169750": ["Рыбий жир", "Козье молоко"], "5584218651": ["Льняное семя"],
    "5584264918": ["Молочные продукты"],

    # ---- Day 8 (2026-09-02) ----
    "5658764220": ["Клюква", "Черника"],
    "5670474832": ["Зелёная чечевица", "Юкка Шидигера"],
    "5670477242": ["Клюква", "Цикорий", "Рыбий жир"],
    "5670498757": ["Брокколи", "Юкка Шидигера"],
    "5670499419": ["Клюква", "Цикорий"],
    "5670501608": ["Семена льна", "Клюква", "Цикорий"],
    "5670469697": ["Брокколи", "Сельдерей", "Семена льна"],
    "5664781811": ["Фукус", "Юкка", "Черника"],
    "5664784502": ["Фукус", "Юкка", "Черника"],
    "5670714887": ["Витамин Е"],
    "5632615106": ["Кацуобуси"],
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

    # ---- Day 5 (2026-08-04) ----
    "5345830359": None,  # Belgian recipe — рецептура, не завод; юрлицо не проверено
    "5358300410": "ООО «Ганеша» (Россия, Воронежская обл., г. Борисоглебск)",
    "5357288481": "Италия",
    "5358242931": None, "5358240926": None,  # Зоопторг — вероятно дистрибьютор, не завод
    "5359892602": None, "5359453152": None, "5359660535": None,
    "5359822345": None,
    "5357563517": None, "5357853793": None,
    "5331257302": None,
    "5358325770": None,
    "5358242441": None, "5358240663": None, "5358241683": None,
    "5358240549": None,  # бренд АВЗ узнаваем, но юрлицо/адрес не подтверждены отдельно

    # ---- Day 6 (2026-08-12) ----
    "5430117217": None, "5423933879": None, "5423741999": None,
    "5432594014": None,  # "Произведено на Алтае" — регион, не юрлицо
    "5419467431": None, "5418435027": None, "5417460160": None,
    "5432515655": None, "5431773123": None, "5384510562": None,
    "5429960930": "Германия",
    "5430728946": None, "5430727673": "Франция", "5430194418": "Германия",
    "5424780704": "Италия", "5424781094": "Италия", "5424780086": "Италия",
    "5424781029": "Италия", "5424783286": "Италия", "5424776640": "Италия",
    "5424785573": "Италия", "5424783287": "Италия", "5424784781": "Италия",
    "5424782869": "Италия",
    "5428395063": None, "5428925488": None,  # "Сделано в России" заявлено, юрлицо не читается
    "5429259395": None, "5431102190": None,
    "5430211436": None,  # не проверено отдельно для этого SKU
    "5432604231": None, "5432823007": None,

    # ---- Доразбор 12.08 ----
    "5433103013": None, "5433102619": None, "5433103250": None,
    "5433103558": None, "5433104303": None, "5433102991": None,
    "5433102865": None, "5433103171": None, "5436654518": None,
    "5424786354": "Италия", "5424780417": "Италия",

    # ---- Day 7 (2026-08-24): Каша для собак ----
    "5542503462": None, "5445105149": None, "5445127064": None,
    "5456129061": None, "5456129661": None, "5456131380": None, "5456131513": None,
    "5479813617": None, "5500647058": None,

    # ---- Day 7 (2026-08-24): Сухой корм для собак ----
    "5515836917": None, "5515838796": None, "5524575556": None,
    "5529408722": None, "5530070606": None, "5555438805": None,
    "5542850495": None, "5556742789": None, "5566163855": None,
    "5566490497": None, "5568844559": "Италия", "5574603661": None,
    "5578347424": None,

    # ---- Day 7 (2026-08-24): Влажный корм для собак ----
    "5471075781": None, "5530369991": None, "5530650366": None, "5530650986": None,
    "5544614194": "Австрия", "5549300374": None, "5556733970": None, "5556739255": None,

    # ---- Day 7 (2026-08-24): Лакомства для собак ----
    "5544395456": None, "5555142584": None, "5555380815": None, "5555663607": None,
    "5556221688": None, "5556538078": None, "5556653006": None, "5577356638": None,
    "5577356979": None, "5577357024": None, "5577357304": None, "5577357359": None,
    "5577358728": None, "5577359119": None, "5577402177": None, "5585499483": None,

    # ---- Day 7 (2026-08-24): Сухой корм для кошек ----
    "5515840508": None, "5515844102": None, "5542531112": None, "5542594928": None,
    "5543934520": None, "5543935300": None, "5543935895": None, "5543936157": None,
    "5543936716": None, "5554107233": "Италия", "5554109198": "Италия",
    "5554109696": "Италия", "5555801791": None, "5556095810": "Италия",
    "5556862888": None, "5566319732": None, "5566555324": None, "5566602205": None,
    "5566632364": None, "5566657335": None, "5568745202": None,
    "5568879450": "Испания", "5574288140": None, "5578339795": None,

    # ---- Day 7 (2026-08-24): Влажный корм для кошек ----
    "5484409205": "Франция", "5484413359": "Франция", "5484418520": "Франция",
    "5530649635": None, "5530649765": None, "5530650056": None, "5530367150": None,
    "5530650988": None, "5530651154": None, "5530651642": None, "5530653845": None,
    "5530654016": None, "5538615147": None, "5543025974": None, "5544350676": None,
    "5544357753": None, "5551551944": None, "5566574938": None, "5566598587": None,
    "5566620215": None, "5566642607": None,

    # ---- Day 7 (2026-08-24): Лакомства для кошек ----
    "5542724663": None, "5543308856": None, "5543534231": None, "5543551335": None,
    "5549365822": None, "5579052755": None, "5579058668": None, "5579066625": None,
    "5582640835": None, "5584169750": None, "5584218651": None, "5584264918": None,

    # ---- Day 8 (2026-09-02) ----
    "5658764220": None,
    "5670474832": None, "5670477242": None, "5670498757": None, "5670499419": None, "5670501608": None,
    "5669273502": None, "5670094001": None, "5670712472": None, "5670714813": None,
    "5670718836": None, "5670720625": None, "5671464738": None, "5671505639": None,
    "5623144132": "Farmina Pet Foods, Италия", "5634101309": None,
    "5662623115": None, "5662623376": None, "5662626493": None,
    "5670255238": None, "5670469697": None,
    "5654017462": None, "5664781811": None, "5664784502": None, "5668947557": None, "5669095176": None,
    "5670228940": None, "5670237889": None, "5670251820": None, "5670328613": None, "5670392893": None,
    "5670714887": None, "5670716904": None, "5670719609": None, "5670721156": None, "5671557401": None,
    "5632614860": None, "5632615231": None, "5632624539": None, "5654589093": None, "5654591896": None,
    "5654592771": None, "5654599026": None, "5654599608": None,
    "5670711483": None, "5670715414": None, "5670717416": None, "5670718664": None, "5670720986": None,
    "5671459555": None, "5671490575": None, "5671509904": None, "5671561038": None,
    "5632614370": None, "5632614943": None, "5632615075": None, "5632615106": None, "5632619730": None,
    "5654586643": None, "5654587557": None, "5654590710": None, "5654591217": None, "5654591653": None,
    "5654592626": None, "5654594462": None, "5654596008": None, "5654598009": None, "5654598442": None,
    "5654598890": None, "5654599052": None, "5654599091": None, "5654599124": None,
    "5670711038": None, "5671503941": None,
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
ws4["A1"] = "Аналитика: популярность направлений клеймов среди новых карточек"
ws4["A1"].font = title_font
ws4["A2"] = (f"Выборка: {len(products)} карточек (собаки+кошки, корма и лакомства) за 7 снимков — "
             "11 отобраны 2026-07-31 через сортировку «Новинки» (не подтверждено ID), "
             "13 добавлены 2026-08-01, 27 добавлены 2026-08-03, 17 добавлены 2026-08-04, 42 добавлены 2026-08-12 "
             "(включая 11, найденных доп. аудитом ?sorting=new — пропущены при первом проходе), 103 добавлены "
             "2026-08-24 и 74 добавлены 2026-09-02 (9-дневный разрыв со снимка от 24.08) — все "
             f"{len(products)} подтверждены сравнением ID с базой (объективно новые). Полностью разобраны "
             "(клеймы+протеин+ингредиенты+завод) все объективно новые карточки, кроме 230 дублей по весу/размеру "
             "упаковки той же рецептуры или повторного листинга того же товара другим продавцом (не считаются "
             "повторно ни по одному признаку) и 1 сопутствующего товара (трава для кошек — набор для выращивания, "
             "не корм/лакомство).")
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
                          f"{protein_known} SKU, где протеин назван на лицевой стороне упаковки, из {PROTEIN_TOTAL_SKUS} SKU всего "
                          f"(все объективно новые SKU за 7 снимков, 2026-07-31 — 2026-09-02)")
next_row = add_tag_table(ws4, next_row + 3, "Ингредиенты-клеймы (лицевая сторона упаковки) — популярность среди новых карточек",
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
chart.title = "Новых карточек с данной категорией клейма (из {})".format(n_products)
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
ws4.cell(row=concl_row, column=1, value="Выводы (наблюдения: 2026-07-31 — 2026-09-02)").font = Font(name=FONT_NAME, bold=True, size=12)
conclusions = [
    "1. Натуральность состава (NAT) остаётся на первом месте, но разрыв с функциональной пользой (FUNC) резко "
    "сократился — 128 из 287 карточек (45%) против 116 (40%). Впервые за наблюдения FUNC почти сравнялась с NAT "
    "как самое частое направление клеймов.",
    "2. НОВОЕ: вкусовая привлекательность (TASTE) практически с нуля выросла до 38 из 287 (13%) и вошла в топ-7 "
    "категорий — почти целиком благодаря одномоментному запуску японского бренда INABA (CIAO/Churu/Yaki для "
    "собак и кошек), у которого большинство SKU без функционального заявления позиционируются именно через вкус "
    "и формат лакомства, а не через пользу для здоровья.",
    "3. САМЫЙ КРУПНЫЙ БРЕНД-ЗАПУСК за всё время наблюдений: INABA единомоментно вывела около 27 отдельных SKU "
    "сразу в 4 подкатегориях (лакомства для собак и кошек, с побочными находками в кормах) — линейки Churu "
    "(лакомство-пюре, более 15 вкусо-функциональных вариаций), Yaki Sasami/Yaki Miksu (запечённые снеки), "
    "Tobitsuku, CIAO Paste и CIAO Pon Churu суфле. Это превосходит даже предыдущий рекорд (AlphaPet/Wanpy, "
    "день 7) и подтверждает вывод №6 предыдущего отчёта: один бренд-зонтик способен единомоментно создать "
    "десятки почти неразличимых карточек.",
    "4. AWARD расширил линейку сухих кормов для кошек сразу 5 SKU (Hairball & Indoor, Urinary, Indoor Big cats, "
    "Sterilized x2) — точное зеркало уже известного шаблона AWARD для собак (100% BIO, Nutraceutical Formula, "
    "Superpremium, Quality for Champions), адаптированного под кошачьи функциональные ниши.",
    "5. ПОДТВЕРЖДЕНО повторно и в наиболее чистом виде: подкатегория «Каша для собак» за 9-дневный разрыв дала "
    "24 кандидата с ID выше порога — и все 24 (100%) оказались повторными листингами уже известной 6-вкусовой "
    "линейки АЛЬМАКОРМ (это уже третий цикл переиздания той же линейки за время наблюдений). Из 112 кандидатов "
    "по всем 7 подкатегориям в сумме лишь 74 (66%) оказались настоящими новыми карточками — прямое численное "
    "продолжение вывода №7 предыдущего отчёта об эффекте 3P-продаж.",
    "6. Мнямс продолжает самую широкую диверсификацию форматов среди всех брендов в выборке: за один день добавил "
    "8 сухих лакомств для собак (джерки, дентал-снеки, рубец, трахея, колбаски), 8 новых консервов/паучей для "
    "собак (муссы, паштеты, террин) и 6 позиций для кошек (линейка Ферма кота Фёдора, крем-лакомства) — бренд "
    "теперь присутствует практически во всех типах продукции трекера одновременно.",
    "7. Завод-изготовитель: доля Италии среди идентифицированных производителей осталась на прежнем уровне "
    "(22 из 44, 50%) — новых итальянских брендов в этом снимке не нашлось, доля просто перестала расти. Для "
    "подавляющего большинства новых брендов дня (INABA, Мнямс, Wanpy-конкуренты P.E.P.P.O. и Вкусная Миска, "
    "AlphaPet WOW) завод не подтверждён — японские и российские мелкие бренды систематически хуже раскрывают "
    "производственную цепочку, чем европейские премиальные.",
    "8. Протеиновый расклад: курица укрепила лидерство (78 из 250 определённых, 31%, было 27%) — драйвер тот же "
    "INABA, где курица — основной протеин почти всех сухих снеков и очень многих Churu-вариаций. Индейка (44, "
    "18%) и говядина (40, 16%) сохранили 2-3 места.",
    "9. Ветеринарная диета (VET) впервые представлена сразу двумя разными форматами позиционирования в одном "
    "снимке: Royal Canin Expert Neutered («№1 по рекомендации ветеринарных врачей») и GRANDORF Veterinary Diet "
    "Hypoallergenic («Approved by Vets») — оба явно апеллируют к врачебному одобрению как отдельному "
    "доказательству качества, а не просто описывают диетическую функцию.",
    "10. Клеймы внутри линейки INABA демонстрируют новый, не встречавшийся ранее паттерн: один и тот же базовый "
    "продукт (лакомство-пюре Churu, курица) выпускается в 4+ параллельных вариантах, различающихся ТОЛЬКО "
    "заявленной функцией на этикетке (кожа/шерсть, ЖКТ, pH мочи, мочеполовая система) при идентичном составе и "
    "формате — то есть бренд буквально мультиплицирует SKU через маркетинговую сегментацию по функции, а не "
    "через реальную вариацию продукта.",
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
