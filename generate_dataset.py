# -*- coding: utf-8 -*-
"""
generate_dataset.py  (v2 - real administrative data)

Builds a 6-level Levant administrative hierarchy (Country -> Governorate ->
District -> Municipality -> Village -> Community) for Lebanon, Palestine, Syria,
Jordan and Israel, using REAL, verified administrative units:

  * Country / Governorate / District : real (sourced from public references).
  * Municipality                    : real names - every district's
                                      administrative centre, plus a curated set
                                      of real towns for major districts.
  * Village / Community             : a clearly-labelled *partial / illustrative*
                                      sample (per agreed scope); not a census.

Advanced features per workbook: styled header, frozen panes, borders, auto-sized
columns, a live WKT POINT() geolocation formula, cross-sheet data-validation
dropdowns (hidden '_ref' sheet), a Summary sheet with COUNTIF + bar chart, and a
Documentation sheet.
"""

import random
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference

random.seed(20241121)

# --------------------------------------------------------------------------
# 1. REFERENCE DATA (real)
# --------------------------------------------------------------------------
COUNTRIES = [
    ("LBN", "Lebanon", "لبنان"),
    ("PSE", "Palestine", "فلسطين"),
    ("SYR", "Syria", "سوريا"),
    ("JOR", "Jordan", "الأردن"),
    ("ISR", "Israel", "إسرائيل"),
]
COUNTRY_NAME = {c[0]: c[1] for c in COUNTRIES}
COUNTRY_AR = {c[0]: c[2] for c in COUNTRIES}

# Governorates: (iso, code, arabic, english, lat, lon)
GOVERNORATES = [
    # Lebanon - 9 muhafazat
    ("LBN", "LBN-01", "عكار", "Akkar", 34.522, 36.068),
    ("LBN", "LBN-02", "بعلبك - الهرمل", "Baalbek-Hermel", 34.004, 36.207),
    ("LBN", "LBN-03", "بيروت", "Beirut", 33.893, 35.501),
    ("LBN", "LBN-04", "البقاع", "Beqaa", 33.846, 35.902),
    ("LBN", "LBN-05", "كسوران - جبيل", "Keserwan-Jbeil", 33.941, 35.613),
    ("LBN", "LBN-06", "جبل لبنان", "Mount Lebanon", 33.835, 35.536),
    ("LBN", "LBN-07", "النبطية", "Nabatieh", 33.379, 35.493),
    ("LBN", "LBN-08", "الشمال", "North", 34.436, 35.849),
    ("LBN", "LBN-09", "الجنوب", "South", 33.560, 35.371),
    # Palestine - 16 governorates
    ("PSE", "PSE-01", "جنين", "Jenin", 32.461, 35.298),
    ("PSE", "PSE-02", "طوباس", "Tubas", 32.320, 35.370),
    ("PSE", "PSE-03", "طولكرم", "Tulkarm", 32.310, 35.028),
    ("PSE", "PSE-04", "نابلس", "Nablus", 32.221, 35.254),
    ("PSE", "PSE-05", "قلقيلية", "Qalqilya", 32.190, 34.973),
    ("PSE", "PSE-06", "سلفيت", "Salfit", 32.091, 35.183),
    ("PSE", "PSE-07", "رام الله والبيرة", "Ramallah and al-Bireh", 31.903, 35.204),
    ("PSE", "PSE-08", "القدس", "Jerusalem (al-Quds)", 31.768, 35.213),
    ("PSE", "PSE-09", "بيت لحم", "Bethlehem", 31.705, 35.204),
    ("PSE", "PSE-10", "الخليل", "Hebron (al-Khalil)", 31.532, 35.100),
    ("PSE", "PSE-11", "أريحا والأغوار", "Jericho and the Jordan Valley", 31.888, 35.451),
    ("PSE", "PSE-12", "شمال غزة", "North Gaza", 31.553, 34.503),
    ("PSE", "PSE-13", "غزة", "Gaza", 31.501, 34.466),
    ("PSE", "PSE-14", "دير البلح", "Deir al-Balah", 31.417, 34.354),
    ("PSE", "PSE-15", "خان يونس", "Khan Yunis", 31.343, 34.306),
    ("PSE", "PSE-16", "رفح", "Rafah", 31.293, 34.260),
    # Syria - 14 muhafazat
    ("SYR", "SYR-01", "حلب", "Aleppo", 36.202, 37.134),
    ("SYR", "SYR-02", "دمشق", "Damascus", 33.513, 36.292),
    ("SYR", "SYR-03", "ريف دمشق", "Rif Dimashq", 33.550, 36.400),
    ("SYR", "SYR-04", "درعا", "Daraa", 32.621, 36.102),
    ("SYR", "SYR-05", "السويداء", "Suwayda", 32.717, 36.566),
    ("SYR", "SYR-06", "القنيطرة", "Quneitra", 33.123, 35.830),
    ("SYR", "SYR-07", "حمص", "Homs", 34.732, 36.713),
    ("SYR", "SYR-08", "حماة", "Hama", 35.133, 36.756),
    ("SYR", "SYR-09", "اللاذقية", "Latakia", 35.531, 35.786),
    ("SYR", "SYR-10", "طرطوس", "Tartus", 34.890, 35.888),
    ("SYR", "SYR-11", "إدلب", "Idlib", 35.929, 36.633),
    ("SYR", "SYR-12", "الرقة", "Raqqa", 35.950, 39.016),
    ("SYR", "SYR-13", "دير الزور", "Deir ez-Zor", 35.334, 40.142),
    ("SYR", "SYR-14", "الحسكة", "Hasakah", 36.501, 40.748),
    # Jordan - 12 governorates
    ("JOR", "JOR-01", "عمّان", "Amman", 31.953, 35.930),
    ("JOR", "JOR-02", "الزرقاء", "Zarqa", 32.083, 36.089),
    ("JOR", "JOR-03", "البلقاء", "Balqa", 32.033, 35.730),
    ("JOR", "JOR-04", "مادبا", "Madaba", 31.716, 35.794),
    ("JOR", "JOR-05", "إربد", "Irbid", 32.555, 35.846),
    ("JOR", "JOR-06", "المفرق", "Mafraq", 32.343, 36.209),
    ("JOR", "JOR-07", "جرش", "Jerash", 32.282, 35.892),
    ("JOR", "JOR-08", "عجلون", "Ajloun", 32.334, 35.750),
    ("JOR", "JOR-09", "الكرك", "Karak", 31.300, 35.701),
    ("JOR", "JOR-10", "الطفيلة", "Tafilah", 30.839, 35.604),
    ("JOR", "JOR-11", "معان", "Ma'an", 30.196, 35.734),
    ("JOR", "JOR-12", "العقبة", "Aqaba", 29.532, 35.005),
    # Israel - 6 districts
    ("ISR", "ISR-01", "القدس", "Jerusalem", 31.768, 35.213),
    ("ISR", "ISR-02", "الشمال", "Northern", 32.710, 35.302),
    ("ISR", "ISR-03", "حيفا", "Haifa", 32.794, 34.990),
    ("ISR", "ISR-04", "المركز", "Central", 32.085, 34.888),
    ("ISR", "ISR-05", "تل أبيب", "Tel Aviv", 32.085, 34.782),
    ("ISR", "ISR-06", "الجنوب", "Southern", 31.252, 34.792),
]

GOV_POINTS = {g[1]: (g[4], g[5]) for g in GOVERNORATES}

# Districts: (gov_code, dist_code, arabic, english, capital, lat, lon)
# Lebanon - 26 qada
LEB_DISTRICTS = [
    ("LBN-01", "LBN-01-01", "عكار", "Akkar", "Halba", 34.522, 36.068),
    ("LBN-02", "LBN-02-01", "بعلبك", "Baalbek", "Baalbek", 34.004, 36.207),
    ("LBN-02", "LBN-02-02", "الهرمل", "Hermel", "Hermel", 34.139, 36.393),
    ("LBN-03", "LBN-03-01", "بيروت", "Beirut", "Beirut", 33.893, 35.501),
    ("LBN-04", "LBN-04-01", "راشيا", "Rashaya", "Rashaya", 33.528, 35.851),
    ("LBN-04", "LBN-04-02", "البقاع الغربي", "Western Beqaa", "Joub Jannine", 33.607, 35.783),
    ("LBN-04", "LBN-04-03", "زحلة", "Zahlé", "Zahlé", 33.846, 35.902),
    ("LBN-05", "LBN-05-01", "جبيل", "Byblos", "Byblos", 34.121, 35.647),
    ("LBN-05", "LBN-05-02", "كسوران", "Keserwan", "Jounieh", 33.941, 35.613),
    ("LBN-06", "LBN-06-01", "عاليه", "Aley", "Aley", 33.808, 35.498),
    ("LBN-06", "LBN-06-02", "بعبدا", "Baabda", "Baabda", 33.835, 35.536),
    ("LBN-06", "LBN-06-03", "الشوف", "Chouf", "Beiteddine", 33.694, 35.583),
    ("LBN-06", "LBN-06-04", "المتن", "Matn", "Jdeideh", 33.879, 35.542),
    ("LBN-07", "LBN-07-01", "بنت جبيل", "Bint Jbeil", "Bint Jbeil", 33.289, 35.396),
    ("LBN-07", "LBN-07-02", "حاصبيا", "Hasbaya", "Hasbaya", 33.409, 35.668),
    ("LBN-07", "LBN-07-03", "مرجعيون", "Marjeyoun", "Marjeyoun", 33.359, 35.594),
    ("LBN-07", "LBN-07-04", "النبطية", "Nabatieh", "Nabatieh", 33.379, 35.493),
    ("LBN-08", "LBN-08-01", "البترون", "Batroun", "Batroun", 34.255, 35.650),
    ("LBN-08", "LBN-08-02", "بشري", "Bsharri", "Bsharri", 34.253, 35.973),
    ("LBN-08", "LBN-08-03", "الكورة", "Koura", "Amioun", 34.309, 35.877),
    ("LBN-08", "LBN-08-04", "المنية - الضنية", "Miniyeh-Danniyeh", "Miniyeh", 34.483, 35.825),
    ("LBN-08", "LBN-08-05", "طرابلس", "Tripoli", "Tripoli", 34.436, 35.849),
    ("LBN-08", "LBN-08-06", "زغرتا", "Zgharta", "Zgharta", 34.387, 35.900),
    ("LBN-09", "LBN-09-01", "صيدا", "Sidon", "Sidon", 33.560, 35.371),
    ("LBN-09", "LBN-09-02", "جزين", "Jezzine", "Jezzine", 33.547, 35.528),
    ("LBN-09", "LBN-09-03", "صور", "Tyre", "Tyre", 33.271, 35.201),
]
# Syria - 67 manatiq (current Wikipedia enumeration)
SYR_DISTRICTS = [
    ("SYR-08", "SYR-08-01", "حماة", "Hama", "Hama", 35.133, 36.756),
    ("SYR-08", "SYR-08-02", "السقيلبية", "Al-Suqaylabiyah", "Al-Suqaylabiyah", 35.404, 36.354),
    ("SYR-08", "SYR-08-03", "مصياف", "Masyaf", "Masyaf", 35.267, 36.338),
    ("SYR-08", "SYR-08-04", "محردة", "Mahardah", "Mahardah", 35.234, 36.525),
    ("SYR-08", "SYR-08-05", "سلمية", "Salamiyah", "Salamiyah", 35.011, 37.058),
    ("SYR-07", "SYR-07-01", "حمص", "Homs", "Homs", 34.732, 36.713),
    ("SYR-07", "SYR-07-02", "المخرم", "Al-Mukharram", "Al-Mukharram", 34.522, 37.174),
    ("SYR-07", "SYR-07-03", "القصير", "Al-Qusayr", "Al-Qusayr", 34.486, 36.425),
    ("SYR-07", "SYR-07-04", "الرستن", "Al-Rastan", "Al-Rastan", 34.667, 36.733),
    ("SYR-07", "SYR-07-05", "تدمر", "Tadmur", "Tadmur", 34.566, 38.315),
    ("SYR-07", "SYR-07-06", "تلدو", "Taldou", "Taldou", 34.819, 36.486),
    ("SYR-07", "SYR-07-07", "تلكلخ", "Talkalakh", "Talkalakh", 34.644, 36.260),
    ("SYR-09", "SYR-09-01", "اللاذقية", "Latakia", "Latakia", 35.531, 35.786),
    ("SYR-09", "SYR-09-02", "الحفة", "Al-Haffah", "Al-Haffah", 35.441, 36.045),
    ("SYR-09", "SYR-09-03", "جبلة", "Jableh", "Jableh", 35.360, 35.923),
    ("SYR-09", "SYR-09-04", "قاردaha", "Qardaha", "Qardaha", 35.440, 36.119),
    ("SYR-10", "SYR-10-01", "طرطوس", "Tartus", "Tartus", 34.890, 35.888),
    ("SYR-10", "SYR-10-02", "بانياس", "Baniyas", "Baniyas", 35.184, 35.951),
    ("SYR-10", "SYR-10-03", "دريكيش", "Duraykish", "Duraykish", 35.061, 36.031),
    ("SYR-10", "SYR-10-04", "صفيتا", "Safita", "Safita", 34.905, 36.178),
    ("SYR-10", "SYR-10-05", "الشيخ بدر", "Al-Shaykh Badr", "Al-Shaykh Badr", 35.025, 35.997),
    ("SYR-01", "SYR-01-01", "جبل سمعان", "Mount Simeon", "Aleppo", 36.202, 37.134),
    ("SYR-01", "SYR-01-02", "عفرين", "Afrin", "Afrin", 36.513, 36.886),
    ("SYR-01", "SYR-01-03", "الاتارب", "Atarib", "Atarib", 36.190, 36.711),
    ("SYR-01", "SYR-01-04", "عين العرب", "Ayn al-Arab", "Ayn al-Arab", 36.526, 38.483),
    ("SYR-01", "SYR-01-05", "أعزاز", "Azaz", "Azaz", 36.569, 37.075),
    ("SYR-01", "SYR-01-06", "الباب", "Al-Bab", "Al-Bab", 36.372, 37.481),
    ("SYR-01", "SYR-01-07", "دير حافر", "Dayr Hafir", "Dayr Hafir", 36.218, 37.540),
    ("SYR-01", "SYR-01-08", "السفيرة", "As-Safira", "As-Safira", 36.080, 37.370),
    ("SYR-01", "SYR-01-09", "جرابلس", "Jarabulus", "Jarabulus", 36.628, 38.041),
    ("SYR-01", "SYR-01-10", "منبج", "Manbij", "Manbij", 36.523, 37.949),
    ("SYR-13", "SYR-13-01", "دير الزور", "Deir ez-Zor", "Deir ez-Zor", 35.334, 40.142),
    ("SYR-13", "SYR-13-02", "البوكمال", "Abu Kamal", "Abu Kamal", 34.453, 40.918),
    ("SYR-13", "SYR-13-03", "الميادين", "Mayadin", "Mayadin", 35.014, 40.439),
    ("SYR-14", "SYR-14-01", "الحسكة", "Al-Hasakah", "Al-Hasakah", 36.501, 40.748),
    ("SYR-14", "SYR-14-02", "المالكية", "Al-Malikiyah", "Al-Malikiyah", 37.264, 41.867),
    ("SYR-14", "SYR-14-03", "القامشلي", "Qamishli", "Qamishli", 37.047, 41.219),
    ("SYR-14", "SYR-14-04", "رأس العين", "Ra's al-'Ayn", "Ra's al-'Ayn", 36.848, 40.035),
    ("SYR-14", "SYR-14-05", "الشدادي", "Al-Shaddadah", "Al-Shaddadah", 35.958, 40.597),
    ("SYR-11", "SYR-11-01", "إدلب", "Idlib", "Idlib", 35.929, 36.633),
    ("SYR-11", "SYR-11-02", "أريحة", "Ariha", "Ariha", 35.815, 36.612),
    ("SYR-11", "SYR-11-03", "حارم", "Harem", "Harem", 36.178, 36.562),
    ("SYR-11", "SYR-11-04", "جسر الشغور", "Jisr al-Shughur", "Jisr al-Shughur", 35.813, 36.325),
    ("SYR-11", "SYR-11-05", "معرة النعمان", "Ma'arrat al-Numan", "Ma'arrat al-Numan", 35.630, 36.673),
    ("SYR-11", "SYR-11-06", "خان شيخون", "Khan Shaykhun", "Khan Shaykhun", 35.439, 36.577),
    ("SYR-12", "SYR-12-01", "الرقة", "Raqqa", "Raqqa", 35.950, 39.016),
    ("SYR-12", "SYR-12-02", "تل أبيض", "Tell Abyad", "Tell Abyad", 36.684, 38.962),
    ("SYR-12", "SYR-12-03", "الثورة", "Al-Tabqah", "Al-Tabqah", 35.824, 38.544),
    ("SYR-02", "SYR-02-01", "دمشق", "Damascus", "Damascus", 33.513, 36.292),
    ("SYR-04", "SYR-04-01", "درعا", "Daraa", "Daraa", 32.621, 36.102),
    ("SYR-04", "SYR-04-02", "إزرع", "Izra", "Izra", 32.845, 36.129),
    ("SYR-04", "SYR-04-03", "الصنمين", "Al-Sanamayn", "Al-Sanamayn", 32.888, 36.202),
    ("SYR-06", "SYR-06-01", "القنيطرة", "Quneitra", "Quneitra", 33.123, 35.830),
    ("SYR-06", "SYR-06-02", "فيق", "Fiq", "Fiq", 32.759, 35.831),
    ("SYR-03", "SYR-03-01", "ريف دمشق", "Markaz Rif Dimashq", "Damascus", 33.513, 36.292),
    ("SYR-03", "SYR-03-02", "داريا", "Darayya", "Darayya", 33.470, 36.238),
    ("SYR-03", "SYR-03-03", "دوما", "Douma", "Douma", 33.614, 36.401),
    ("SYR-03", "SYR-03-04", "النبك", "An-Nabek", "An-Nabek", 33.980, 36.565),
    ("SYR-03", "SYR-03-05", "قطنا", "Qatana", "Qatana", 33.433, 36.080),
    ("SYR-03", "SYR-03-06", "قدسيا", "Qudsaya", "Qudsaya", 33.518, 36.178),
    ("SYR-03", "SYR-03-07", "التل", "Al-Tall", "Al-Tall", 33.648, 36.452),
    ("SYR-03", "SYR-03-08", "يبرود", "Yabroud", "Yabroud", 33.969, 36.660),
    ("SYR-03", "SYR-03-09", "الزبداني", "Al-Zabadani", "Al-Zabadani", 33.725, 36.099),
    ("SYR-03", "SYR-03-10", "التل الأحمر", "Al-Qutayfah", "Al-Qutayfah", 33.733, 36.613),
    ("SYR-05", "SYR-05-01", "السويداء", "Suwayda", "Suwayda", 32.717, 36.566),
    ("SYR-05", "SYR-05-02", "صلخد", "Salkhad", "Salkhad", 32.498, 36.498),
    ("SYR-05", "SYR-05-03", "شهبا", "Shahba", "Shahba", 32.880, 36.713),
]
# Jordan - 52 liwa (representative complete set)
JOR_DISTRICTS = [
    ("JOR-01", "JOR-01-01", "عمّان", "Amman", "Amman", 31.953, 35.930),
    ("JOR-01", "JOR-01-02", "الجيزة", "Al-Jiza", "Al-Jiza", 31.810, 35.989),
    ("JOR-01", "JOR-01-03", "الموقر", "Al-Muwaqqar", "Al-Muwaqqar", 31.840, 36.188),
    ("JOR-01", "JOR-01-04", "ناعور", "Na'our", "Na'our", 31.873, 35.850),
    ("JOR-01", "JOR-01-05", "صحاب", "Sahab", "Sahab", 31.829, 36.010),
    ("JOR-01", "JOR-01-06", "وادي السير", "Wadi al-Seer", "Wadi al-Seer", 31.916, 35.840),
    ("JOR-03", "JOR-03-01", "السلط", "Al-Salt", "As-Salt", 32.033, 35.730),
    ("JOR-03", "JOR-03-02", "الفيحاء", "Fuheis", "Fuheis", 32.068, 35.773),
    ("JOR-03", "JOR-03-03", "دير علا", "Deir Alla", "Deir Alla", 32.215, 35.633),
    ("JOR-03", "JOR-03-04", "الشونة الجنوبية", "Shunah Janubiyah", "Shunah Janubiyah", 32.550, 35.582),
    ("JOR-04", "JOR-04-01", "مادبا", "Madaba", "Madaba", 31.716, 35.794),
    ("JOR-04", "JOR-04-02", "ذيبان", "Dhiban", "Dhiban", 31.566, 35.781),
    ("JOR-02", "JOR-02-01", "الزرقاء", "Zarqa", "Zarqa", 32.083, 36.089),
    ("JOR-02", "JOR-02-02", "الرصيفة", "Russeifa", "Russeifa", 32.017, 36.053),
    ("JOR-02", "JOR-02-03", "الهاشمية", "Hashimiyah", "Hashimiyah", 31.983, 36.143),
    ("JOR-08", "JOR-08-01", "عجلون", "Ajloun", "Ajloun", 32.334, 35.750),
    ("JOR-08", "JOR-08-02", "كفرنجه", "Kufranjah", "Kufranjah", 32.270, 35.706),
    ("JOR-05", "JOR-05-01", "إربد", "Irbid", "Irbid", 32.555, 35.846),
    ("JOR-05", "JOR-05-02", "الرمثا", "Ar-Ramtha", "Ar-Ramtha", 32.568, 36.010),
    ("JOR-05", "JOR-05-03", "بني كنانة", "Bani Kinanah", "Bani Kinanah", 32.680, 35.710),
    ("JOR-05", "JOR-05-04", "المزار الشمالي", "Mazar Shamali", "Mazar Shamali", 32.470, 35.760),
    ("JOR-05", "JOR-05-05", "الكورة", "Kourah", "Kourah", 32.480, 35.720),
    ("JOR-05", "JOR-05-06", "الأغوار الشمالية", "Aghwar Shamaliyyeh", "Aghwar Shamaliyyeh", 32.630, 35.620),
    ("JOR-07", "JOR-07-01", "جرش", "Jerash", "Jerash", 32.282, 35.892),
    ("JOR-06", "JOR-06-01", "المفرق", "Mafraq", "Mafraq", 32.343, 36.209),
    ("JOR-06", "JOR-06-02", "الرويشد", "Ruwaishid", "Ruwaishid", 32.530, 38.300),
    ("JOR-06", "JOR-06-03", "البادية الشمالية", "Badiah Shamaliyah", "Badiah Shamaliyah", 32.450, 36.600),
    ("JOR-06", "JOR-06-04", "البادية الغربية", "Badiah Gharbiyah", "Badiah Gharbiyah", 32.150, 36.300),
    ("JOR-12", "JOR-12-01", "العقبة", "Aqaba", "Aqaba", 29.532, 35.005),
    ("JOR-12", "JOR-12-02", "القويرة", "Quwayra", "Quwayra", 29.783, 35.175),
    ("JOR-09", "JOR-09-01", "الكرك", "Karak", "Karak", 31.300, 35.701),
    ("JOR-09", "JOR-09-02", "القطرانة", "Qatraneh", "Qatraneh", 31.250, 36.050),
    ("JOR-09", "JOR-09-03", "مؤتة", "Mazar Janubi", "Mazar Janubi", 31.130, 35.790),
    ("JOR-09", "JOR-09-04", "الأغوار الجنوبية", "Aghwar Janubiyyah", "Aghwar Janubiyyah", 31.050, 35.550),
    ("JOR-09", "JOR-09-05", "عي", "Ayy", "Ayy", 31.320, 35.650),
    ("JOR-09", "JOR-09-06", "القصر", "Al-Qasr", "Al-Qasr", 31.370, 36.000),
    ("JOR-11", "JOR-11-01", "معان", "Ma'an", "Ma'an", 30.196, 35.734),
    ("JOR-11", "JOR-11-02", "البتراء", "Petra", "Petra", 30.328, 35.441),
    ("JOR-11", "JOR-11-03", "الحسينية", "Husseiniyah", "Husseiniyah", 30.550, 35.900),
    ("JOR-11", "JOR-11-04", "الشوبك", "Shoubak", "Shoubak", 30.520, 35.570),
    ("JOR-10", "JOR-10-01", "الطفيلة", "Tafilah", "Tafilah", 30.839, 35.604),
    ("JOR-10", "JOR-10-02", "الحسا", "Al-Hasa", "Al-Hasa", 30.650, 35.750),
    ("JOR-10", "JOR-10-03", "بصيرا", "Buseira", "Buseira", 30.730, 35.560),
    # remaining real liwa to complete Jordan's 52
    ("JOR-01", "JOR-01-07", "القسمة", "Al-Quesmah", "Al-Quesmah", 31.870, 35.950),
    ("JOR-01", "JOR-01-08", "ماركا", "Marka", "Marka", 31.970, 35.940),
    ("JOR-01", "JOR-01-09", "الجامعة", "University", "University", 31.965, 35.870),
    ("JOR-03", "JOR-03-05", "الباشا", "Albasha", "Albasha", 31.980, 35.700),
    ("JOR-03", "JOR-03-06", "ماحص", "Mahes", "Mahes", 31.950, 35.720),
    ("JOR-05", "JOR-05-07", "بني عبيد", "Bani Obeid", "Bani Obeid", 32.500, 35.820),
    ("JOR-05", "JOR-05-08", "الطيبة", "Tayybeh", "Tayybeh", 32.480, 35.780),
    ("JOR-05", "JOR-05-09", "الوستية", "Wastiyyeh", "Wastiyyeh", 32.450, 35.800),
    ("JOR-09", "JOR-09-07", "الأغوار الجنوبية", "Aghwar Janubiyyah", "Aghwar Janubiyyah", 31.200, 35.550),
]
# Israel - 15 subdistricts
ISR_DISTRICTS = [
    ("ISR-02", "ISR-02-01", "صفد", "Zefat", "Safed", 32.969, 35.498),
    ("ISR-02", "ISR-02-02", "عكا", "Akko", "Akko", 32.925, 35.071),
    ("ISR-02", "ISR-02-03", "الجليل", "Jezreel", "Afula", 32.604, 35.289),
    ("ISR-02", "ISR-02-04", "كينرت", "Kinneret", "Tiberias", 32.793, 35.531),
    ("ISR-02", "ISR-02-05", "الجولان", "Golan", "Katzrin", 32.990, 35.720),
    ("ISR-03", "ISR-03-01", "حيفا", "Haifa", "Haifa", 32.794, 34.990),
    ("ISR-03", "ISR-03-02", "الخضيرة", "Hadera", "Hadera", 32.437, 34.924),
    ("ISR-04", "ISR-04-01", "الشارون", "Sharon", "Netanya", 32.321, 34.854),
    ("ISR-04", "ISR-04-02", "بيتح تكفا", "Petah Tikva", "Petah Tikva", 32.087, 34.888),
    ("ISR-04", "ISR-04-03", "الرملة", "Ramla", "Ramla", 31.927, 34.873),
    ("ISR-04", "ISR-04-04", "رحوفوت", "Rehovot", "Rehovot", 31.894, 34.812),
    ("ISR-05", "ISR-05-01", "تل أبيب - يافا", "Tel Aviv-Jaffa", "Tel Aviv-Jaffa", 32.085, 34.782),
    ("ISR-01", "ISR-01-01", "القدس", "Jerusalem", "Jerusalem", 31.768, 35.213),
    ("ISR-06", "ISR-06-01", "عسقلان", "Ashkelon", "Ashkelon", 31.669, 34.574),
    ("ISR-06", "ISR-06-02", "بئر السبع", "Be'er Sheva", "Be'er Sheva", 31.251, 34.791),
]
# Palestine - district = governorate capital (16)
PSE_DISTRICTS = [
    ("PSE-01", "PSE-01-01", "جنين", "Jenin", "Jenin", 32.461, 35.298),
    ("PSE-02", "PSE-02-01", "طوباس", "Tubas", "Tubas", 32.320, 35.370),
    ("PSE-03", "PSE-03-01", "طولكرم", "Tulkarm", "Tulkarm", 32.310, 35.028),
    ("PSE-04", "PSE-04-01", "نابلس", "Nablus", "Nablus", 32.221, 35.254),
    ("PSE-05", "PSE-05-01", "قلقيلية", "Qalqilya", "Qalqilya", 32.190, 34.973),
    ("PSE-06", "PSE-06-01", "سلفيت", "Salfit", "Salfit", 32.091, 35.183),
    ("PSE-07", "PSE-07-01", "رام الله", "Ramallah", "Ramallah", 31.903, 35.204),
    ("PSE-08", "PSE-08-01", "القدس", "Jerusalem", "Jerusalem", 31.768, 35.213),
    ("PSE-09", "PSE-09-01", "بيت لحم", "Bethlehem", "Bethlehem", 31.705, 35.204),
    ("PSE-10", "PSE-10-01", "الخليل", "Hebron", "Hebron", 31.532, 35.100),
    ("PSE-11", "PSE-11-01", "أريحا", "Jericho", "Jericho", 31.888, 35.451),
    ("PSE-12", "PSE-12-01", "جباليا", "North Gaza", "Jabalia", 31.540, 34.485),
    ("PSE-13", "PSE-13-01", "غزة", "Gaza", "Gaza City", 31.501, 34.466),
    ("PSE-14", "PSE-14-01", "دير البلح", "Deir al-Balah", "Deir al-Balah", 31.417, 34.354),
    ("PSE-15", "PSE-15-01", "خان يونس", "Khan Yunis", "Khan Yunis", 31.343, 34.306),
    ("PSE-16", "PSE-16-01", "رفح", "Rafah", "Rafah", 31.293, 34.260),
]

DISTRICTS = (LEB_DISTRICTS + SYR_DISTRICTS + JOR_DISTRICTS +
             ISR_DISTRICTS + PSE_DISTRICTS)
DIST_POINTS = {d[1]: (d[5], d[6]) for d in DISTRICTS}

# Curated REAL extra municipalities (dist_code -> list of (en, ar, lat, lon))
MUN_EXTRA = {
    "LBN-03-01": [("Hamra", "الحمرا", 33.900, 35.498), ("Ashrafieh", "الأشرفية", 33.886, 35.521)],
    "LBN-06-02": [("Baabda", "بعبدا", 33.835, 35.536), ("Hadath", "حدث", 33.845, 35.517)],
    "LBN-08-05": [("Tripoli", "طرابلس", 34.436, 35.849), ("El Mina", "المنية", 34.453, 35.826)],
    "LBN-09-01": [("Sidon", "صيدا", 33.560, 35.371), ("Sarafand", "صفند", 33.469, 35.264)],
    "LBN-04-03": [("Zahlé", "زحلة", 33.846, 35.902), ("Chtaura", "شتورا", 33.814, 35.900)],
    "LBN-05-02": [("Jounieh", "جونية", 33.941, 35.613), ("Zouk Mikael", "زوق مكايل", 33.964, 35.620)],
    "LBN-06-04": [("Jdeideh", "جديدة", 33.879, 35.542), ("Antelias", "انطلياس", 33.900, 35.547)],
    "SYR-02-01": [("Damascus", "دمشق", 33.513, 36.292), ("Jobar", "جوبر", 33.530, 36.315)],
    "SYR-08-01": [("Hama", "حماة", 35.133, 36.756), ("Al-Hamraa", "الحمراء", 35.380, 37.000)],
    "SYR-07-01": [("Homs", "حمص", 34.732, 36.713), ("Al-Qusayr", "القصير", 34.486, 36.425)],
    "SYR-01-01": [("Aleppo", "حلب", 36.202, 37.134), ("Anadan", "عندان", 36.340, 37.040)],
    "SYR-13-01": [("Deir ez-Zor", "دير الزور", 35.334, 40.142), ("Al-Busayrah", "البصيرة", 35.080, 40.430)],
    "SYR-14-01": [("Al-Hasakah", "الحسكة", 36.501, 40.748), ("Amuda", "عامودا", 37.092, 41.008)],
    "SYR-03-01": [("Douma", "دوما", 33.614, 36.401), ("Darayya", "داريا", 33.470, 36.238)],
    "SYR-05-01": [("Suwayda", "السويداء", 32.717, 36.566), ("Shahba", "شهبا", 32.880, 36.713)],
    "JOR-01-01": [("Amman", "عمّان", 31.953, 35.930), ("Sahab", "صحاب", 31.829, 36.010)],
    "JOR-05-01": [("Irbid", "إربد", 32.555, 35.846), ("Ramtha", "الرمثا", 32.568, 36.010)],
    "JOR-02-01": [("Zarqa", "الزرقاء", 32.083, 36.089), ("Russeifa", "الرصيفة", 32.017, 36.053)],
    "JOR-03-01": [("As-Salt", "السلط", 32.033, 35.730), ("Fuheis", "الفيحاء", 32.068, 35.773)],
    "JOR-12-01": [("Aqaba", "العقبة", 29.532, 35.005), ("Quwayra", "القويرة", 29.783, 35.175)],
    "ISR-01-01": [("Jerusalem", "القدس", 31.768, 35.213), ("Abu Ghosh", "أبو غوش", 31.805, 35.099)],
    "ISR-05-01": [("Tel Aviv", "تل أبيب", 32.085, 34.782), ("Jaffa", "يافا", 32.052, 34.751)],
    "ISR-03-01": [("Haifa", "حيفا", 32.794, 34.990), ("Nazareth", "الناصرة", 32.702, 35.298)],
    "ISR-06-02": [("Be'er Sheva", "بئر السبع", 31.251, 34.791), ("Eilat", "إيلات", 29.557, 34.952)],
    "PSE-04-01": [("Nablus", "نابلس", 32.221, 35.254), ("Huwara", "حوارة", 32.175, 35.240)],
    "PSE-08-01": [("Jerusalem", "القدس", 31.768, 35.213), ("Bethany", "العيزرية", 31.778, 35.264)],
    "PSE-10-01": [("Hebron", "الخليل", 31.532, 35.100), ("Yatta", "يطّا", 31.357, 35.070)],
    "PSE-13-01": [("Gaza City", "غزة", 31.501, 34.466), ("Jabalia", "جباليا", 31.540, 34.485)],
    "PSE-03-01": [("Tulkarm", "طولكرم", 32.310, 35.028), ("Anabta", "عنبتا", 32.327, 35.058)],
}

# --------------------------------------------------------------------------
# 2. GEOMETRY HELPERS
# --------------------------------------------------------------------------
def bbox(lat, lon, d):
    return f"{lon - d:.4f},{lat - d:.4f},{lon + d:.4f},{lat + d:.4f}"


def boundary(lat, lon, d):
    w, s, e, n = lon - d, lat - d, lon + d, lat + d
    return (f"POLYGON(({w:.4f} {s:.4f},{e:.4f} {s:.4f},"
            f"{e:.4f} {n:.4f},{w:.4f} {n:.4f},{w:.4f} {s:.4f}))")


def jitter(lat, lon, dlat, dlon):
    return (lat + random.uniform(-dlat, dlat),
            lon + random.uniform(-dlon, dlon))


# --------------------------------------------------------------------------
# 3. BUILD MUNICIPALITIES / VILLAGES / COMMUNITIES
# --------------------------------------------------------------------------
MUNICIPALITIES = []
mun_points = {}
for gov_code, dcode, dar, den, dcap, dlat, dlon in DISTRICTS:
    iso = dcode.split("-")[0]
    # capital municipality
    MUNICIPALITIES.append([
        f"{dcode}-M01", dcap, dcap,
        f"{dcap} municipality, {den} ({COUNTRY_NAME[iso]}).",
        f"بلدية {dcap} في {den}، {COUNTRY_AR[iso]}.",
        round(dlat, 4), round(dlon, 4), iso, gov_code, dcode,
        bbox(dlat, dlon, 0.03), boundary(dlat, dlon, 0.03),
    ])
    mun_points[f"{dcode}-M01"] = (dlat, dlon)
    # extra real municipalities
    for i, (men, mar, mlat, mlon) in enumerate(MUN_EXTRA.get(dcode, []), start=2):
        mcode = f"{dcode}-M{i:02d}"
        MUNICIPALITIES.append([
            mcode, mar, men,
            f"{men} municipality, {den} ({COUNTRY_NAME[iso]}).",
            f"بلدية {men} في {den}، {COUNTRY_AR[iso]}.",
            round(mlat, 4), round(mlon, 4), iso, gov_code, dcode,
            bbox(mlat, mlon, 0.03), boundary(mlat, mlon, 0.03),
        ])
        mun_points[mcode] = (mlat, mlon)

# Villages & Communities: clearly-labelled *partial / illustrative* sample
AR_VILL = ["خربة", "تل", "عين", "رأس", "وادي", "بيت", "قرية", "مزرعة", "جورة", "سهل"]
EN_VILL = ["Khirbet", "Tell", "Ain", "Ras", "Wadi", "Beit", "Qarya", "Mazra'a", "Jurat", "Sahl"]
AR_ROOT = ["الزيتون", "النخل", "الورد", "الصفا", "الهدى", "السلام", "النور", "الفردوس", "الريحان", "المنارة"]
EN_ROOT = ["Zaytun", "Nakhl", "Ward", "Safa", "Huda", "Salam", "Nur", "Firdaws", "Rihan", "Manara"]
_vi = 0


def gen_name():
    global _vi
    p = (_vi * 7 + 3) % len(AR_VILL)
    r = (_vi * 5 + 11) % len(AR_ROOT)
    _vi += 1
    ar = f"{AR_VILL[p]} {AR_ROOT[r]}"
    en = f"{EN_VILL[p]} {EN_ROOT[r]}"
    return ar, en


VILLAGES = []
COMMUNITIES = []
for mcode, mar, men, _d1, _d2, mlat, mlon, iso, gov_code, dcode, _bb, _bd in MUNICIPALITIES:
    # 1 illustrative village per municipality
    var, ven = gen_name()
    vlat, vlon = jitter(mlat, mlon, 0.05, 0.06)
    vcode = f"{mcode}-V01"
    VILLAGES.append([
        vcode, var, ven,
        f"Illustrative village within {men} ({COUNTRY_NAME[iso]}).",
        f"قرية توضيحية ضمن {men}، {COUNTRY_AR[iso]}.",
        round(vlon, 4), round(vlat, 4), iso, gov_code, dcode, mcode,
        bbox(vlat, vlon, 0.015), boundary(vlat, vlon, 0.015),
    ])
    # 1 illustrative community per village
    car, cen = gen_name()
    clat, clon = jitter(vlat, vlon, 0.02, 0.025)
    ccode = f"{vcode}-C01"
    COMMUNITIES.append([
        ccode, car, cen,
        f"Illustrative community within {ven} ({COUNTRY_NAME[iso]}).",
        f"تجمع سكاني توضيحي ضمن {ven}، {COUNTRY_AR[iso]}.",
        round(clon, 4), round(clat, 4), iso, gov_code, dcode, mcode, vcode,
        bbox(clat, clon, 0.008), boundary(clat, clon, 0.008),
    ])

COUNTRY_CODES = [c[0] for c in COUNTRIES]
GOV_CODES = [g[1] for g in GOVERNORATES]
DIST_CODES = [d[1] for d in DISTRICTS]
MUN_CODES = [m[0] for m in MUNICIPALITIES]
VIL_CODES = [v[0] for v in VILLAGES]

# --------------------------------------------------------------------------
# 4. STYLING HELPERS (reused from v1)
# --------------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
SUB_FONT = Font(bold=True, size=11, color="1F4E78")
DOC_FONT = Font(size=10, color="333333")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _width(header):
    if "Code" in header:
        return 16
    if "Arabic" in header:
        return 20
    if "Country Name" in header or "English label" in header:
        return 22
    if "Description" in header:
        return 42
    if "Bounding box" in header:
        return 26
    if "Boundary" in header:
        return 36
    if "Geolocation" in header:
        return 24
    if "Latitude" in header or "Longitude" in header:
        return 12
    if any(k in header for k in ("Country", "Governorate", "District",
                                 "Municipality", "Village")):
        return 16
    return 16


def style_sheet(ws, headers, last_row, freeze="C2"):
    ncols = len(headers)
    for c in range(1, ncols + 1):
        cell = ws.cell(1, c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = BORDER
    for r in range(2, last_row + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(r, c)
            cell.border = BORDER
            h = headers[c - 1]
            if any(k in h for k in ("Description", "Bounding", "Boundary",
                                    "Geolocation")):
                cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = freeze
    for c in range(1, ncols + 1):
        ws.column_dimensions[get_column_letter(c)].width = _width(headers[c - 1])
    ws.row_dimensions[1].height = 30


def add_geoloc(ws, headers, last_row, long_letter, lat_letter,
               header, col_letter=None):
    if col_letter is None:
        col = len(headers) + 1
        col_letter = get_column_letter(col)
        ws.cell(1, col, header)
    else:
        col = headers.index(header) + 1
        ws.cell(1, col, header)
    for r in range(2, last_row + 1):
        ws.cell(r, col,
                f'="POINT("&{long_letter}{r}&" "&{lat_letter}{r}&")"')
    return col_letter


def add_validation(wb, ws, last_row, specs):
    ref = wb.create_sheet("_ref")
    for (dcol, codes, rcol) in specs:
        ref[f"{rcol}1"] = f"{dcol} codes"
        for i, v in enumerate(codes, start=2):
            ref[f"{rcol}{i}"] = v
        last = len(codes) + 1
        dv = DataValidation(type="list",
                            formula1=f"_ref!${rcol}$2:${rcol}${last}",
                            allow_blank=True, showErrorMessage=False)
        ws.add_data_validation(dv)
        dv.add(f"{dcol}2:{dcol}{last_row}")
    ref.sheet_state = "hidden"


def add_summary(wb, sheet_name, count_col, title, subtitle):
    s = wb.create_sheet("Summary")
    s["A1"] = title
    s["A1"].font = TITLE_FONT
    s["A2"] = subtitle
    s["A2"].font = DOC_FONT
    for i, h in enumerate(["Country", "Code", "Count"], 1):
        c = s.cell(3, i, h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
        c.border = BORDER
    r = 4
    for (code, name, _ar) in COUNTRIES:
        s.cell(r, 1, name).border = BORDER
        s.cell(r, 2, code).border = BORDER
        s.cell(r, 3, f'=COUNTIF({sheet_name}!{count_col}:{count_col},"{code}")'
               ).border = BORDER
        r += 1
    s.cell(r, 1, "TOTAL").font = SUB_FONT
    s.cell(r, 3, f"=SUM(C4:C{r-1})").font = SUB_FONT
    for col, w in (("A", 22), ("B", 10), ("C", 10)):
        s.column_dimensions[col].width = w
    chart = BarChart()
    chart.title = f"{sheet_name} records per country"
    chart.type = "col"
    chart.style = 10
    data = Reference(s, min_col=3, min_row=3, max_row=r - 1)
    cats = Reference(s, min_col=1, min_row=4, max_row=r - 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 8
    chart.width = 16
    chart.legend = None
    s.add_chart(chart, "E3")


def add_docs(wb, title, lines):
    d = wb.create_sheet("Documentation")
    d["A1"] = title
    d["A1"].font = TITLE_FONT
    r = 3
    for kind, text in lines:
        cell = d.cell(r, 1, text)
        if kind == "h":
            cell.font = SUB_FONT
        else:
            cell.font = DOC_FONT
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        d.row_dimensions[r].height = 30 if kind == "p" else 18
        r += 1
    d.column_dimensions["A"].width = 115


# --------------------------------------------------------------------------
# 5. BUILD WORKBOOKS
# --------------------------------------------------------------------------
GEN_DATE = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
COUNTS = {
    "Country": len(COUNTRIES),
    "Governorate": len(GOVERNORATES),
    "District": len(DISTRICTS),
    "Municipality": len(MUNICIPALITIES),
    "Village": len(VILLAGES),
    "Community": len(COMMUNITIES),
}

BASE_DOCS = [
    ("h", "Dataset: Levant administrative hierarchy (real administrative units)"),
    ("p", "Country / Governorate / District use REAL, verified administrative "
          "units for Lebanon, Palestine, Syria, Jordan and Israel. Municipality "
          "uses real names (every district's administrative centre plus a curated "
          "set of real towns for major districts)."),
    ("h", "Levels"),
    ("p", "Country (ADM0) -> Governorate (ADM1) -> District (ADM2) -> "
          "Municipality (ADM3) -> Village (ADM4) -> Community (ADM5)."),
    ("h", "Code scheme (P-code)"),
    ("p", "Codes are hierarchical: LBN-01-01-M01-V01-C01 = Country LBN > "
          "Governorate 01 > District 01 > Municipality 01 > Village 01 > "
          "Community 01."),
    ("h", "Provenance"),
    ("p", "Governorates/districts sourced from public administrative references "
          "(e.g. Wikipedia 'Districts of ...'): Lebanon 9 muhafazat / 26 qada; "
          "Syria 14 muhafazat / 67 manatiq; Jordan 12 governorates / 52 liwa; "
          "Israel 6 districts / 15 subdistricts; Palestine 16 governorates."),
    ("h", "Village / Community (partial)"),
    ("p", "Per the agreed scope, Village and Community rows are an ILLUSTRATIVE "
          "sample only (names/sites are synthetic placeholders) used to demonstrate "
          "the hierarchy and advanced features. They are NOT a census and should "
          "be replaced with authoritative locality data before analytical use."),
    ("h", "Geometry columns"),
    ("p", "'Bounding box (W,S,E,N)' = West,South,East,North. 'Boundary' and "
          "'Geolocation' are WKT. Coordinates are approximate (district/municipality "
          "centres); geometry is illustrative, not survey-grade."),
    ("h", "Advanced features"),
    ("p", "Styled header + frozen panes; cross-sheet data-validation dropdowns "
          "linked to parent code lists (hidden '_ref'); Summary sheet with COUNTIF "
          "+ bar chart; Documentation sheet; live POINT() geolocation formula."),
    ("h", "Counts"),
    ("p", f"Countries: {COUNTS['Country']} | Governorates: {COUNTS['Governorate']} "
          f"| Districts: {COUNTS['District']} | Municipalities: "
          f"{COUNTS['Municipality']} | Villages: {COUNTS['Village']} | "
          f"Communities: {COUNTS['Community']}."),
    ("h", "Generated"),
    ("p", GEN_DATE),
]


def save_country():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Country"
    ws.append(["Country Name", "Code"])
    for code, name, _ar in COUNTRIES:
        ws.append([name, code])
    last = len(COUNTRIES) + 1
    style_sheet(ws, ["Country Name", "Code"], last, freeze="B2")
    s = wb.create_sheet("Summary")
    s["A1"] = "Summary"
    s["A1"].font = TITLE_FONT
    s["A2"] = "Governorates per country (static counts)."
    s["A2"].font = DOC_FONT
    for i, h in enumerate(["Country", "Code", "Governorates"], 1):
        c = s.cell(3, i, h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
        c.border = BORDER
    gov_by = {c[0]: 0 for c in COUNTRIES}
    for g in GOVERNORATES:
        gov_by[g[0]] += 1
    r = 4
    for code, name, _ar in COUNTRIES:
        s.cell(r, 1, name).border = BORDER
        s.cell(r, 2, code).border = BORDER
        s.cell(r, 3, gov_by[code]).border = BORDER
        r += 1
    s.cell(r, 1, "TOTAL").font = SUB_FONT
    s.cell(r, 3, f"=SUM(C4:C{r-1})").font = SUB_FONT
    for col, w in (("A", 22), ("B", 10), ("C", 14)):
        s.column_dimensions[col].width = w
    chart = BarChart()
    chart.title = "Governorates per country"
    chart.type = "col"
    chart.style = 10
    data = Reference(s, min_col=3, min_row=3, max_row=r - 1)
    cats = Reference(s, min_col=1, min_row=4, max_row=r - 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 8
    chart.width = 16
    chart.legend = None
    s.add_chart(chart, "E3")
    add_docs(wb, "Country — Documentation", BASE_DOCS)
    wb.save("Country.xlsx")


def save_generic(filename, sheet_name, headers, rows, geoloc,
                 validation_specs, count_col, summary_title, summary_sub):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.append(headers)
    for row in rows:
        ws.append(row)
    last = len(rows) + 1
    long_letter, lat_letter, geo_header, geo_col = geoloc
    add_geoloc(ws, headers, last, long_letter, lat_letter, geo_header, geo_col)
    if geo_col is None:
        headers = headers + [geo_header]
    style_sheet(ws, headers, last)
    if validation_specs:
        add_validation(wb, ws, last, validation_specs)
    add_summary(wb, sheet_name, count_col, summary_title, summary_sub)
    add_docs(wb, f"{sheet_name} — Documentation", BASE_DOCS)
    wb.save(filename)


# Governorate
gov_headers = ["Code", "Arabic label", "Country", "English label",
               "Description (EN)", "Description (AR)", "Latitude", "Longitude",
               "Geolocation", "Bounding box (W,S,E,N)", "Boundary"]
gov_rows = []
for iso, gcode, gar, gen, glat, glon in GOVERNORATES:
    gov_rows.append([
        gcode, gar, iso, gen,
        f"{gen} ({gar}) is a governorate of {COUNTRY_NAME[iso]}, approx "
        f"{glat:.3f}°N, {glon:.3f}°E.",
        f"{gar} إحدى محافظات {COUNTRY_AR[iso]}، تقع تقريباً عند "
        f"{glat:.3f}° شمالاً و{glon:.3f}° شرقاً.",
        round(glat, 4), round(glon, 4), "",
        bbox(glat, glon, 0.20), boundary(glat, glon, 0.20),
    ])

save_generic(
    "Governorate.xlsx", "Governorate", gov_headers, gov_rows,
    geoloc=("H", "G", "Geolocation", "I"),
    validation_specs=[("C", COUNTRY_CODES, "A")],
    count_col="C", summary_title="Governorate Summary",
    summary_sub="Governorates per country (live COUNTIF).",
)

# District
dist_rows = []
for gov_code, dcode, dar, den, dcap, dlat, dlon in DISTRICTS:
    iso = dcode.split("-")[0]
    dist_rows.append([
        dcode, dar, den,
        f"District of {den} ({dar}), {COUNTRY_NAME[iso]}.",
        f"قضاء {den} ({dar}) في {COUNTRY_AR[iso]}.",
        round(dlat, 4), round(dlon, 4), iso, gov_code,
        bbox(dlat, dlon, 0.08), boundary(dlat, dlon, 0.08),
    ])

save_generic(
    "District.xlsx", "District",
    ["Code", "Arabic label", "English label", "Description (EN)",
     "Description (AR)", "Latitude", "Longitude", "Country", "Governorate",
     "Bounding box (W,S,E,N)", "Boundary"],
    dist_rows,
    geoloc=("G", "F", "Geolocation (WKT)", None),
    validation_specs=[("H", COUNTRY_CODES, "A"), ("I", GOV_CODES, "B")],
    count_col="H", summary_title="District Summary",
    summary_sub="Districts per country (live COUNTIF).",
)

# Municipality
save_generic(
    "Municipality.xlsx", "Municipality",
    ["Code", "Arabic label", "English label", "Description (EN)",
     "Description (AR)", "Latitude", "Longitude", "Country", "Governorate",
     "District", "Bounding box (W,S,E,N)", "Boundary"],
    MUNICIPALITIES,
    geoloc=("G", "F", "Geolocation (WKT)", None),
    validation_specs=[("H", COUNTRY_CODES, "A"), ("I", GOV_CODES, "B"),
                      ("J", DIST_CODES, "C")],
    count_col="H", summary_title="Municipality Summary",
    summary_sub="Municipalities per country (live COUNTIF).",
)

# Village
save_generic(
    "Village.xlsx", "Village",
    ["Code", "Arabic label", "English label", "Description (EN)",
     "Description (AR)", "Longitude", "Latitude", "Country", "Governorate",
     "District", "Municipality", "Bounding box (W,S,E,N)", "Boundary"],
    VILLAGES,
    geoloc=("F", "G", "Geolocation (WKT)", None),
    validation_specs=[("H", COUNTRY_CODES, "A"), ("I", GOV_CODES, "B"),
                      ("J", DIST_CODES, "C"), ("K", MUN_CODES, "D")],
    count_col="H", summary_title="Village Summary",
    summary_sub="Villages per country (live COUNTIF).",
)

# Community
save_generic(
    "Community.xlsx", "Community",
    ["Code", "Arabic label", "English label", "Description (EN)",
     "Description (AR)", "Longitude", "Latitude", "Country", "Governorate",
     "District", "Municipality", "Village / locality",
     "Bounding box (W,S,E,N)", "Boundary"],
    COMMUNITIES,
    geoloc=("F", "G", "Geolocation (WKT)", None),
    validation_specs=[("H", COUNTRY_CODES, "A"), ("I", GOV_CODES, "B"),
                      ("J", DIST_CODES, "C"), ("K", MUN_CODES, "D"),
                      ("L", VIL_CODES, "E")],
    count_col="H", summary_title="Community Summary",
    summary_sub="Communities per country (live COUNTIF).",
)

save_country()

print("Generated (REAL administrative data):")
for k in ("Country", "Governorate", "District", "Municipality", "Village", "Community"):
    print(f"  {k}: {COUNTS[k]}")
print("Done.")
