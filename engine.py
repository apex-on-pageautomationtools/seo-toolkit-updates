"""
engine.py - Hardened, human-like Google SERP engine.

This module is the anti-CAPTCHA core. It replaces the old "hit
/search?q=...&num=10&start=0&pws=0" pattern (which Google 403-blocks on sight)
with a human browsing session:

  1.  Deep stealth Chrome (undetected-chromedriver + extra CDP patches).
  2.  Warm-up on the Google homepage (cookies + consent) before any query.
  3.  Query is TYPED into the real search box, char-by-char, then submitted —
      so navigation carries a real referer + form tokens, like a person.
  4.  Deeper results come from human scrolling + the "More results" button
      (continuous scroll), NOT from forged start= URLs.
  5.  Block detection -> exponential backoff -> proxy rotation -> fresh
      identity -> (optional) Buster -> manual-solve pause, in that order.

Nothing here is Google-specific magic; it is ordinary stealth + patience,
which is what keeps the soft-block / reCAPTCHA rate low.
"""

from __future__ import annotations

# Python 3.12 removed distutils -- inject setuptools' version so
# undetected_chromedriver's "from distutils.version import LooseVersion" works
try:
    import setuptools._distutils as _dt
    import sys as _sys
    import types as _types
    if 'distutils' not in _sys.modules:
        _sys.modules['distutils'] = _dt
    for _k, _v in vars(_dt).items():
        if isinstance(_v, _types.ModuleType):
            _subkey = f"distutils.{_k}"
            if _subkey not in _sys.modules:
                _sys.modules[_subkey] = _v
    del _dt, _k, _v, _types, _subkey
except Exception:
    pass

import os
import re
import json
import time
import random
import tempfile
import zipfile
import logging

log = logging.getLogger("grc.engine")

# --------------------------------------------------------------------------- #
# Country -> Google ccTLD + a representative timezone (for emulation)
# --------------------------------------------------------------------------- #
GOOGLE_DOMAINS = {
    "us": "google.com", "gb": "google.co.uk", "au": "google.com.au",
    "ca": "google.ca", "in": "google.co.in", "de": "google.de",
    "fr": "google.fr", "es": "google.es", "it": "google.it",
    "nl": "google.nl", "br": "google.com.br", "mx": "google.com.mx",
    "jp": "google.co.jp", "kr": "google.co.kr", "ru": "google.ru",
    "ae": "google.ae", "sg": "google.com.sg", "za": "google.co.za",
    "nz": "google.co.nz", "ie": "google.ie", "se": "google.se",
    "ch": "google.ch", "pl": "google.pl", "id": "google.co.id",
    "ph": "google.com.ph", "my": "google.com.my", "pk": "google.com.pk",
    "ng": "google.com.ng", "sa": "google.com.sa", "tr": "google.com.tr",
}

COUNTRY_TZ = {
    "us": "America/New_York", "gb": "Europe/London", "au": "Australia/Sydney",
    "ca": "America/Toronto", "in": "Asia/Kolkata", "de": "Europe/Berlin",
    "fr": "Europe/Paris", "es": "Europe/Madrid", "it": "Europe/Rome",
    "nl": "Europe/Amsterdam", "br": "America/Sao_Paulo", "mx": "America/Mexico_City",
    "jp": "Asia/Tokyo", "kr": "Asia/Seoul", "ru": "Europe/Moscow",
    "ae": "Asia/Dubai", "sg": "Asia/Singapore", "za": "Africa/Johannesburg",
    "nz": "Pacific/Auckland", "ie": "Europe/Dublin", "se": "Europe/Stockholm",
    "ch": "Europe/Zurich", "pl": "Europe/Warsaw", "id": "Asia/Jakarta",
    "ph": "Asia/Manila", "my": "Asia/Kuala_Lumpur", "pk": "Asia/Karachi",
    "ng": "Africa/Lagos", "sa": "Asia/Riyadh", "tr": "Europe/Istanbul",
}


COUNTRY_GEO = {
    "us": (40.7128, -74.0060), "gb": (51.5074, -0.1278), "au": (-33.8688, 151.2093),
    "ca": (43.6532, -79.3832), "in": (28.6139, 77.2090), "de": (52.5200, 13.4050),
    "fr": (48.8566, 2.3522), "es": (40.4168, -3.7038), "it": (41.9028, 12.4964),
    "nl": (52.3676, 4.9041), "br": (-23.5505, -46.6333), "mx": (19.4326, -99.1332),
    "jp": (35.6762, 139.6503), "kr": (37.5665, 126.9780), "ru": (55.7558, 37.6173),
    "ae": (25.2048, 55.2708), "sg": (1.3521, 103.8198), "za": (-33.9249, 18.4241),
    "nz": (-36.8485, 174.7633), "ie": (53.3498, -6.2603), "se": (59.3293, 18.0686),
    "ch": (47.3769, 8.5417), "pl": (52.2297, 21.0122), "id": (-6.2088, 106.8456),
    "ph": (14.5995, 120.9842), "my": (3.1390, 101.6869), "pk": (33.6844, 73.0479),
    "ng": (6.5244, 3.3792), "sa": (24.7136, 46.6753), "tr": (41.0082, 28.9784),
}

# City coordinates for CDP geolocation override (Sensors emulation)
CITY_COORDS = {
    # US
    "New York, US": (40.7128, -74.0060), "Los Angeles, US": (34.0522, -118.2437),
    "Chicago, US": (41.8781, -87.6298), "Houston, US": (29.7604, -95.3698),
    "Miami, US": (25.7617, -80.1918), "San Francisco, US": (37.7749, -122.4194),
    "Seattle, US": (47.6062, -122.3321), "Dallas, US": (32.7767, -96.7970),
    "Denver, US": (39.7392, -104.9903), "Boston, US": (42.3601, -71.0589),
    "Atlanta, US": (33.7490, -84.3880), "Phoenix, US": (33.4484, -112.0740),
    "Philadelphia, US": (39.9526, -75.1652), "San Diego, US": (32.7157, -117.1611),
    "Austin, US": (30.2672, -97.7431), "Las Vegas, US": (36.1699, -115.1398),
    "Portland, US": (45.5152, -122.6784), "Nashville, US": (36.1627, -86.7816),
    "Charlotte, US": (35.2271, -80.8431), "San Antonio, US": (29.4241, -98.4936),
    "Columbus, US": (39.9612, -82.9988), "Indianapolis, US": (39.7684, -86.1581),
    "Jacksonville, US": (30.3322, -81.6557), "Fort Worth, US": (32.7555, -97.3308),
    "San Jose, US": (37.3382, -121.8863), "Detroit, US": (42.3314, -83.0458),
    "Memphis, US": (35.1495, -90.0490), "Baltimore, US": (39.2904, -76.6122),
    "Milwaukee, US": (43.0389, -87.9065), "Tucson, US": (32.2226, -110.9747),
    "Sacramento, US": (38.5816, -121.4944), "Kansas City, US": (39.0997, -94.5786),
    "Raleigh, US": (35.7796, -78.6382), "Tampa, US": (27.9506, -82.4572),
    "Minneapolis, US": (44.9778, -93.2650), "Orlando, US": (28.5383, -81.3792),
    "Pittsburgh, US": (40.4406, -79.9959), "Cincinnati, US": (39.1031, -84.5120),
    "New Orleans, US": (29.9511, -90.0715), "Salt Lake City, US": (40.7608, -111.8910),
    "Honolulu, US": (21.3069, -157.8583), "Anchorage, US": (61.2181, -149.9003),
    "St. Louis, US": (38.6270, -90.1994), "Cleveland, US": (41.4993, -81.6944),
    "Virginia Beach, US": (36.8529, -75.9780), "Omaha, US": (41.2565, -95.9345),
    "Oakland, US": (37.8044, -122.2712), "Tulsa, US": (36.1540, -95.9928),
    "Bakersfield, US": (35.3733, -119.0187), "Albuquerque, US": (35.0844, -106.6504),
    # AU
    "Sydney, AU": (-33.8688, 151.2093), "Melbourne, AU": (-37.8136, 144.9631),
    "Brisbane, AU": (-27.4698, 153.0251), "Perth, AU": (-31.9505, 115.8605),
    "Adelaide, AU": (-34.9285, 138.6007), "Gold Coast, AU": (-28.0167, 153.4000),
    "Canberra, AU": (-35.2809, 149.1300), "Newcastle, AU": (-32.9283, 151.7817),
    "Hobart, AU": (-42.8821, 147.3272), "Darwin, AU": (-12.4634, 130.8456),
    "Wollongong, AU": (-34.4278, 150.8931), "Cairns, AU": (-16.9186, 145.7781),
    "Townsville, AU": (-19.2590, 146.8169), "Geelong, AU": (-38.1499, 144.3617),
    "Sunshine Coast, AU": (-26.6500, 153.0667), "Toowoomba, AU": (-27.5598, 151.9507),
    "Ballarat, AU": (-37.5622, 143.8503), "Bendigo, AU": (-36.7570, 144.2794),
    "Launceston, AU": (-41.4332, 147.1441), "Rockhampton, AU": (-23.3791, 150.5100),
    # CA
    "Toronto, CA": (43.6532, -79.3832), "Vancouver, CA": (49.2827, -123.1207),
    "Montreal, CA": (45.5017, -73.5673), "Calgary, CA": (51.0447, -114.0719),
    "Ottawa, CA": (45.4215, -75.6972), "Edmonton, CA": (53.5461, -113.4938),
    "Winnipeg, CA": (49.8951, -97.1384), "Quebec City, CA": (46.8139, -71.2080),
    "Hamilton, CA": (43.2557, -79.8711), "Victoria, CA": (48.4284, -123.3656),
    "Halifax, CA": (44.6488, -63.5752), "Saskatoon, CA": (52.1332, -106.6700),
    "Regina, CA": (50.4452, -104.6189), "Mississauga, CA": (43.5890, -79.6441),
    "Brampton, CA": (43.7315, -79.7624), "Kitchener, CA": (43.4516, -80.4925),
    "London, CA": (42.9849, -81.2453), "Surrey, CA": (49.1913, -122.8490),
    "Markham, CA": (43.8561, -79.3370), "St. John's, CA": (47.5615, -52.7126),
    "Fredericton, CA": (45.9636, -66.6431), "Charlottetown, CA": (46.2382, -63.1311),
    "Whitehorse, CA": (60.7212, -135.0568), "Yellowknife, CA": (62.4540, -114.3718),
    "Iqaluit, CA": (63.7467, -68.5170),
    # GB
    "London, GB": (51.5074, -0.1278), "Manchester, GB": (53.4808, -2.2426),
    "Birmingham, GB": (52.4862, -1.8904), "Glasgow, GB": (55.8642, -4.2518),
    "Edinburgh, GB": (55.9533, -3.1883), "Liverpool, GB": (53.4084, -2.9916),
    "Leeds, GB": (53.8008, -1.5491), "Bristol, GB": (51.4545, -2.5879),
    "Sheffield, GB": (53.3811, -1.4701), "Cardiff, GB": (51.4816, -3.1791),
    "Belfast, GB": (54.5973, -5.9301), "Newcastle, GB": (54.9783, -1.6178),
    "Nottingham, GB": (52.9548, -1.1581), "Southampton, GB": (50.9097, -1.4044),
    "Cambridge, GB": (52.2053, 0.1218),
    # IN
    "Mumbai, IN": (19.0760, 72.8777), "New Delhi, IN": (28.6139, 77.2090),
    "Bangalore, IN": (12.9716, 77.5946), "Chennai, IN": (13.0827, 80.2707),
    "Hyderabad, IN": (17.3850, 78.4867), "Kolkata, IN": (22.5726, 88.3639),
    "Pune, IN": (18.5204, 73.8567), "Jaipur, IN": (26.9124, 75.7873),
    "Ahmedabad, IN": (23.0225, 72.5714), "Lucknow, IN": (26.8467, 80.9462),
    "Chandigarh, IN": (30.7333, 76.7794), "Indore, IN": (22.7196, 75.8577),
    "Bhopal, IN": (23.2599, 77.4126), "Nagpur, IN": (21.1458, 79.0882),
    "Kochi, IN": (9.9312, 76.2673), "Coimbatore, IN": (11.0168, 76.9558),
    "Surat, IN": (21.1702, 72.8311), "Visakhapatnam, IN": (17.6868, 83.2185),
    "Patna, IN": (25.6093, 85.1376), "Thiruvananthapuram, IN": (8.5241, 76.9366),
    "Itanagar, IN": (27.0844, 93.6053), "Guwahati, IN": (26.1445, 91.7362),
    "Raipur, IN": (21.2514, 81.6296), "Panaji, IN": (15.4909, 73.8278),
    "Gurugram, IN": (28.4595, 77.0266), "Shimla, IN": (31.1048, 77.1734),
    "Ranchi, IN": (23.3441, 85.3096), "Imphal, IN": (24.8170, 93.9368),
    "Shillong, IN": (25.5788, 91.8933), "Aizawl, IN": (23.7271, 92.7176),
    "Kohima, IN": (25.6751, 94.1086), "Bhubaneswar, IN": (20.2961, 85.8245),
    "Amritsar, IN": (31.6340, 74.8723), "Gangtok, IN": (27.3389, 88.6065),
    "Agartala, IN": (23.8315, 91.2868), "Dehradun, IN": (30.3165, 78.0322),
    "Port Blair, IN": (11.6234, 92.7265), "Daman, IN": (20.3974, 72.8328),
    "Srinagar, IN": (34.0837, 74.7973), "Leh, IN": (34.1526, 77.5771),
    "Kavaratti, IN": (10.5669, 72.6420), "Puducherry, IN": (11.9416, 79.8083),
    # DE
    "Berlin, DE": (52.5200, 13.4050), "Munich, DE": (48.1351, 11.5820),
    "Hamburg, DE": (53.5511, 9.9937), "Frankfurt, DE": (50.1109, 8.6821),
    "Cologne, DE": (50.9375, 6.9603), "Stuttgart, DE": (48.7758, 9.1829),
    "Dusseldorf, DE": (51.2277, 6.7735), "Dortmund, DE": (51.5136, 7.4653),
    "Essen, DE": (51.4556, 7.0116), "Leipzig, DE": (51.3397, 12.3731),
    "Potsdam, DE": (52.3906, 13.0645), "Bremen, DE": (53.0793, 8.8017),
    "Hanover, DE": (52.3759, 9.7320), "Rostock, DE": (54.0887, 12.1400),
    "Mainz, DE": (49.9929, 8.2473), "Saarbrucken, DE": (49.2401, 6.9969),
    "Magdeburg, DE": (52.1205, 11.6276), "Kiel, DE": (54.3233, 10.1228),
    "Erfurt, DE": (50.9848, 11.0299),
    # FR
    "Paris, FR": (48.8566, 2.3522), "Lyon, FR": (45.7640, 4.8357),
    "Marseille, FR": (43.2965, 5.3698), "Toulouse, FR": (43.6047, 1.4442),
    "Nice, FR": (43.7102, 7.2620), "Bordeaux, FR": (44.8378, -0.5792),
    "Lille, FR": (50.6292, 3.0573), "Strasbourg, FR": (48.5734, 7.7521),
    "Rouen, FR": (49.4432, 1.0999), "Rennes, FR": (48.1173, -1.6778),
    "Nantes, FR": (47.2184, -1.5536), "Orleans, FR": (47.9029, 1.9093),
    "Dijon, FR": (47.3220, 5.0415), "Ajaccio, FR": (41.9192, 8.7386),
    # ES
    "Madrid, ES": (40.4168, -3.7038), "Barcelona, ES": (41.3851, 2.1734),
    "Valencia, ES": (39.4699, -0.3763), "Seville, ES": (37.3891, -5.9845),
    "Malaga, ES": (36.7213, -4.4217), "Bilbao, ES": (43.2630, -2.9350),
    "Zaragoza, ES": (41.6488, -0.8891), "Murcia, ES": (37.9922, -1.1307), "Palma, ES": (39.5696, 2.6502), "Las Palmas, ES": (28.1235, -15.4363),
    # IT
    "Rome, IT": (41.9028, 12.4964), "Milan, IT": (45.4642, 9.1900),
    "Naples, IT": (40.8518, 14.2681), "Turin, IT": (45.0703, 7.6869),
    "Florence, IT": (43.7696, 11.2558), "Bologna, IT": (44.4949, 11.3426),
    "Venice, IT": (45.4408, 12.3155), "Palermo, IT": (38.1157, 13.3615),
    "Bari, IT": (41.1171, 16.8719), "Catanzaro, IT": (38.9098, 16.5877),
    "Cagliari, IT": (39.2238, 9.1217), "Genoa, IT": (44.4056, 8.9463),
    "Ancona, IT": (43.6158, 13.5189), "Pescara, IT": (42.4643, 14.2142),
    "Trieste, IT": (45.6495, 13.7768), "Trento, IT": (46.0748, 11.1217),
    "Perugia, IT": (43.1122, 12.3888), "Potenza, IT": (40.6404, 15.8054),
    "Campobasso, IT": (41.5603, 14.6627), "Aosta, IT": (45.7372, 7.3157),
    # NL
    "Amsterdam, NL": (52.3676, 4.9041), "Rotterdam, NL": (51.9244, 4.4777),
    "The Hague, NL": (52.0705, 4.3007), "Utrecht, NL": (52.0907, 5.1214),
    "Eindhoven, NL": (51.4416, 5.4697), "Tilburg, NL": (51.5555, 5.0913), "Groningen, NL": (53.2194, 6.5665), "Almere, NL": (52.3508, 5.2647), "Breda, NL": (51.5719, 4.7683), "Nijmegen, NL": (51.8425, 5.8528),
    # BR
    "Sao Paulo, BR": (-23.5505, -46.6333), "Rio de Janeiro, BR": (-22.9068, -43.1729),
    "Brasilia, BR": (-15.7975, -47.8919), "Salvador, BR": (-12.9714, -38.5124),
    "Curitiba, BR": (-25.4284, -49.2733),
    "Fortaleza, BR": (-3.7172, -38.5433), "Belo Horizonte, BR": (-19.9167, -43.9345), "Manaus, BR": (-3.119, -60.0217), "Recife, BR": (-8.0476, -34.877), "Porto Alegre, BR": (-30.0346, -51.2177),
    # MX
    "Mexico City, MX": (19.4326, -99.1332), "Guadalajara, MX": (20.6597, -103.3496),
    "Monterrey, MX": (25.6866, -100.3161), "Cancun, MX": (21.1619, -86.8515),
    "Puebla, MX": (19.0414, -98.2063),
    "Tijuana, MX": (32.5149, -117.0382), "Leon, MX": (21.125, -101.686), "Juarez, MX": (31.6904, -106.4245), "Zapopan, MX": (20.7167, -103.4), "Merida, MX": (20.9674, -89.5926),
    # JP
    "Tokyo, JP": (35.6762, 139.6503), "Osaka, JP": (34.6937, 135.5023),
    "Yokohama, JP": (35.4437, 139.6380), "Nagoya, JP": (35.1815, 136.9066),
    "Kyoto, JP": (35.0116, 135.7681),
    "Sapporo, JP": (43.0618, 141.3545), "Fukuoka, JP": (33.5904, 130.4017), "Kobe, JP": (34.6901, 135.1955), "Kawasaki, JP": (35.5308, 139.7029), "Saitama, JP": (35.8617, 139.6455),
    # KR
    "Seoul, KR": (37.5665, 126.9780), "Busan, KR": (35.1796, 129.0756),
    "Incheon, KR": (37.4563, 126.7052), "Daegu, KR": (35.8714, 128.6014),
    "Daejeon, KR": (36.3504, 127.3845), "Gwangju, KR": (35.1595, 126.8526), "Suwon, KR": (37.2636, 127.0286), "Ulsan, KR": (35.5384, 129.3114), "Changwon, KR": (35.2281, 128.6811), "Goyang, KR": (37.6584, 126.832),
    # RU
    "Moscow, RU": (55.7558, 37.6173), "Saint Petersburg, RU": (59.9343, 30.3351),
    "Novosibirsk, RU": (55.0084, 82.9357),
    "Yekaterinburg, RU": (56.8389, 60.6057), "Kazan, RU": (55.8304, 49.0661), "Nizhny Novgorod, RU": (56.2965, 43.9361), "Chelyabinsk, RU": (55.1644, 61.4368), "Samara, RU": (53.2001, 50.15), "Omsk, RU": (54.9885, 73.3242), "Rostov-on-Don, RU": (47.2357, 39.7015),
    # AE
    "Dubai, AE": (25.2048, 55.2708), "Abu Dhabi, AE": (24.4539, 54.3773),
    "Sharjah, AE": (25.3463, 55.4209),
    "Ajman, AE": (25.4052, 55.5136), "Ras Al Khaimah, AE": (25.7895, 55.9432), "Fujairah, AE": (25.1288, 56.3265), "Umm Al Quwain, AE": (25.5647, 55.5533), "Al Ain, AE": (24.2075, 55.7447),
    # SG
    "Singapore, SG": (1.3521, 103.8198),
    "Jurong, SG": (1.3329, 103.7436), "Tampines, SG": (1.3496, 103.9568), "Woodlands, SG": (1.4382, 103.789),
    # ZA
    "Johannesburg, ZA": (-26.2041, 28.0473), "Cape Town, ZA": (-33.9249, 18.4241),
    "Durban, ZA": (-29.8587, 31.0218),
    "Pretoria, ZA": (-25.7479, 28.2293), "Port Elizabeth, ZA": (-33.9608, 25.6022), "Bloemfontein, ZA": (-29.0852, 26.1596), "East London, ZA": (-33.0153, 27.9116), "Nelspruit, ZA": (-25.4753, 30.9694), "Polokwane, ZA": (-23.9045, 29.4689), "Kimberley, ZA": (-28.7282, 24.7499),
    # NZ
    "Auckland, NZ": (-36.8485, 174.7633), "Wellington, NZ": (-41.2865, 174.7762),
    "Christchurch, NZ": (-43.5321, 172.6362), "Hamilton, NZ": (-37.7870, 175.2793),
    "Tauranga, NZ": (-37.6878, 176.1651), "Napier, NZ": (-39.4928, 176.912), "Dunedin, NZ": (-45.8788, 170.5028), "Palmerston North, NZ": (-40.3523, 175.6082), "Nelson, NZ": (-41.2706, 173.284), "Rotorua, NZ": (-38.1368, 176.2497),
    # IE
    "Dublin, IE": (53.3498, -6.2603), "Cork, IE": (51.8985, -8.4756),
    "Galway, IE": (53.2707, -9.0568),
    "Limerick, IE": (52.6638, -8.6267), "Waterford, IE": (52.2593, -7.1101), "Drogheda, IE": (53.7189, -6.3478), "Kilkenny, IE": (52.6541, -7.2448), "Wexford, IE": (52.3369, -6.4633),
    # SE
    "Stockholm, SE": (59.3293, 18.0686), "Gothenburg, SE": (57.7089, 11.9746),
    "Malmo, SE": (55.6050, 13.0038),
    "Uppsala, SE": (59.8586, 17.6389), "Vasteras, SE": (59.6099, 16.5448), "Orebro, SE": (59.2753, 15.2134), "Linkoping, SE": (58.4108, 15.6214), "Helsingborg, SE": (56.0465, 12.6945), "Jonkoping, SE": (57.7826, 14.1618), "Norrkoping, SE": (58.5877, 16.1924),
    # CH
    "Zurich, CH": (47.3769, 8.5417), "Geneva, CH": (46.2044, 6.1432),
    "Bern, CH": (46.9480, 7.4474),
    "Basel, CH": (47.5596, 7.5886), "Lausanne, CH": (46.5197, 6.6323), "Lucerne, CH": (47.0502, 8.3093), "Winterthur, CH": (47.5001, 8.7241), "St. Gallen, CH": (47.4245, 9.3767), "Lugano, CH": (46.0037, 8.9511), "Biel, CH": (47.1368, 7.2468),
    # PL
    "Warsaw, PL": (52.2297, 21.0122), "Krakow, PL": (50.0647, 19.9450),
    "Wroclaw, PL": (51.1079, 17.0385),
    "Lodz, PL": (51.7592, 19.456), "Poznan, PL": (52.4064, 16.9252), "Gdansk, PL": (54.352, 18.6466), "Szczecin, PL": (53.4285, 14.5528), "Bydgoszcz, PL": (53.1235, 18.0084), "Lublin, PL": (51.2465, 22.5684), "Katowice, PL": (50.2649, 19.0238),
    # ID
    "Jakarta, ID": (-6.2088, 106.8456), "Surabaya, ID": (-7.2575, 112.7521),
    "Bali, ID": (-8.3405, 115.0920), "Bandung, ID": (-6.9175, 107.6191),
    "Medan, ID": (3.5952, 98.6722), "Semarang, ID": (-6.9667, 110.4167), "Makassar, ID": (-5.1477, 119.4327), "Palembang, ID": (-2.9761, 104.7754), "Yogyakarta, ID": (-7.7956, 110.3695), "Denpasar, ID": (-8.6705, 115.2126),
    # PH
    "Manila, PH": (14.5995, 120.9842), "Cebu, PH": (10.3157, 123.8854),
    "Davao, PH": (7.1907, 125.4553),
    "Quezon City, PH": (14.676, 121.0437), "Zamboanga, PH": (6.9214, 122.079), "Antipolo, PH": (14.5878, 121.176), "Taguig, PH": (14.5176, 121.0509), "Cagayan de Oro, PH": (8.4542, 124.6319), "Paranaque, PH": (14.4793, 121.0198), "Makati, PH": (14.5547, 121.0244),
    # MY
    "Kuala Lumpur, MY": (3.1390, 101.6869), "Penang, MY": (5.4164, 100.3327),
    "Johor Bahru, MY": (1.4927, 103.7414),
    "Ipoh, MY": (4.5975, 101.0901), "Shah Alam, MY": (3.0733, 101.5185), "Kuching, MY": (1.5535, 110.3593), "Kota Kinabalu, MY": (5.9804, 116.0735), "Malacca City, MY": (2.1896, 102.2501), "Alor Setar, MY": (6.1248, 100.3678), "Petaling Jaya, MY": (3.1073, 101.6067),
    # PK
    "Karachi, PK": (24.8607, 67.0011), "Lahore, PK": (31.5204, 74.3587),
    "Islamabad, PK": (33.6844, 73.0479), "Rawalpindi, PK": (33.5651, 73.0169),
    "Faisalabad, PK": (31.4504, 73.135), "Multan, PK": (30.1575, 71.5249), "Peshawar, PK": (34.0151, 71.5249), "Quetta, PK": (30.1798, 66.975), "Sialkot, PK": (32.4945, 74.5229), "Gujranwala, PK": (32.1877, 74.1945),
    # NG
    "Lagos, NG": (6.5244, 3.3792), "Abuja, NG": (9.0765, 7.3986),
    "Port Harcourt, NG": (4.8156, 7.0498),
    "Kano, NG": (12.0022, 8.592), "Ibadan, NG": (7.3775, 3.947), "Benin City, NG": (6.335, 5.6037), "Kaduna, NG": (10.5222, 7.4383), "Enugu, NG": (6.5244, 7.5086), "Zaria, NG": (11.0804, 7.7076), "Aba, NG": (5.1066, 7.3667),
    # SA
    "Riyadh, SA": (24.7136, 46.6753), "Jeddah, SA": (21.4858, 39.1925),
    "Mecca, SA": (21.3891, 39.8579), "Dammam, SA": (26.3927, 49.9777),
    "Medina, SA": (24.5247, 39.5692), "Taif, SA": (21.2703, 40.4158), "Tabuk, SA": (28.3998, 36.5715), "Buraidah, SA": (26.326, 43.975), "Khobar, SA": (26.2172, 50.1971), "Abha, SA": (18.2164, 42.5053),
    # TR
    "Istanbul, TR": (41.0082, 28.9784), "Ankara, TR": (39.9334, 32.8597),
    "Izmir, TR": (38.4237, 27.1428), "Antalya, TR": (36.8969, 30.7133),
    "Bursa, TR": (40.1826, 29.0665), "Adana, TR": (37.0, 35.3213), "Gaziantep, TR": (37.0662, 37.3833), "Konya, TR": (37.8746, 32.4932), "Mersin, TR": (36.8, 34.6333), "Kayseri, TR": (38.7312, 35.4787),
    # TH
    "Bangkok, TH": (13.7563, 100.5018), "Chiang Mai, TH": (18.7883, 98.9853),
    "Phuket, TH": (7.8804, 98.3923),
    "Pattaya, TH": (12.9236, 100.8825), "Nonthaburi, TH": (13.8622, 100.5144), "Nakhon Ratchasima, TH": (14.9799, 102.0977), "Khon Kaen, TH": (16.4322, 102.8236), "Udon Thani, TH": (17.4138, 102.787), "Hat Yai, TH": (7.0086, 100.4747), "Rayong, TH": (12.6833, 101.2372),
    # VN
    "Ho Chi Minh City, VN": (10.8231, 106.6297), "Hanoi, VN": (21.0278, 105.8342),
    "Da Nang, VN": (16.0544, 108.2022),
    "Can Tho, VN": (10.0452, 105.7469), "Hai Phong, VN": (20.8449, 106.6881), "Bien Hoa, VN": (10.9447, 106.8243), "Nha Trang, VN": (12.2388, 109.1967), "Hue, VN": (16.4637, 107.5909), "Vung Tau, VN": (10.346, 107.0843), "Quy Nhon, VN": (13.7757, 109.2237),
    # AR
    "Buenos Aires, AR": (-34.6037, -58.3816), "Cordoba, AR": (-31.4201, -64.1888),
    "Rosario, AR": (-32.9468, -60.6393),
    "Mendoza, AR": (-32.8895, -68.8458), "La Plata, AR": (-34.9215, -57.9545), "San Miguel de Tucuman, AR": (-26.8241, -65.2226), "Mar del Plata, AR": (-38.0055, -57.5426), "Salta, AR": (-24.7859, -65.4117), "Santa Fe, AR": (-31.6333, -60.7), "Neuquen, AR": (-38.9516, -68.0591),
    # CO
    "Bogota, CO": (4.7110, -74.0721), "Medellin, CO": (6.2442, -75.5812),
    "Cali, CO": (3.4516, -76.5320),
    "Barranquilla, CO": (10.9639, -74.7964), "Cartagena, CO": (10.391, -75.4794), "Cucuta, CO": (7.8939, -72.5078), "Bucaramanga, CO": (7.1193, -73.1227), "Pereira, CO": (4.8133, -75.6961), "Santa Marta, CO": (11.2408, -74.199), "Ibague, CO": (4.4389, -75.2322),
    # CL
    "Santiago, CL": (-33.4489, -70.6693), "Valparaiso, CL": (-33.0472, -71.6127),
    "Concepcion, CL": (-36.8201, -73.0444), "La Serena, CL": (-29.9027, -71.2519), "Antofagasta, CL": (-23.6509, -70.3975), "Temuco, CL": (-38.7359, -72.5904), "Rancagua, CL": (-34.1708, -70.7444), "Talca, CL": (-35.4264, -71.6554), "Iquique, CL": (-20.2307, -70.1357), "Puerto Montt, CL": (-41.4693, -72.9424),
    # PE
    "Lima, PE": (-12.0464, -77.0428), "Arequipa, PE": (-16.4090, -71.5375),
    "Trujillo, PE": (-8.1116, -79.0288), "Chiclayo, PE": (-6.7714, -79.8409), "Piura, PE": (-5.1945, -80.6328), "Iquitos, PE": (-3.7437, -73.2516), "Cusco, PE": (-13.532, -71.9675), "Chimbote, PE": (-9.0853, -78.5783), "Huancayo, PE": (-12.0651, -75.2049), "Tacna, PE": (-18.0146, -70.2536),
    # EG
    "Cairo, EG": (30.0444, 31.2357), "Alexandria, EG": (31.2001, 29.9187),
    "Giza, EG": (30.0131, 31.2089), "Shubra El Kheima, EG": (30.1286, 31.2422), "Port Said, EG": (31.2565, 32.2841), "Suez, EG": (29.9668, 32.5498), "Luxor, EG": (25.6872, 32.6396), "Mansoura, EG": (31.0409, 31.3785), "Tanta, EG": (30.7865, 31.0004), "Aswan, EG": (24.0889, 32.8998),
    # KE
    "Nairobi, KE": (-1.2921, 36.8219), "Mombasa, KE": (-4.0435, 39.6682),
    "Kisumu, KE": (-0.0917, 34.768), "Nakuru, KE": (-0.3031, 36.08), "Eldoret, KE": (0.5143, 35.2698), "Thika, KE": (-1.0332, 37.0692), "Malindi, KE": (-3.2192, 40.1169),
    # PT
    "Lisbon, PT": (38.7223, -9.1393), "Porto, PT": (41.1579, -8.6291),
    "Vila Nova de Gaia, PT": (41.1239, -8.6118), "Amadora, PT": (38.7536, -9.2302), "Braga, PT": (41.5454, -8.4265), "Coimbra, PT": (40.2033, -8.4103), "Funchal, PT": (32.6669, -16.9241), "Setubal, PT": (38.5244, -8.8882), "Almada, PT": (38.6795, -9.1568), "Aveiro, PT": (40.6443, -8.6455),
    # AT
    "Vienna, AT": (48.2082, 16.3738), "Salzburg, AT": (47.8095, 13.0550),
    "Graz, AT": (47.0707, 15.4395), "Linz, AT": (48.3069, 14.2858), "Innsbruck, AT": (47.2692, 11.4041), "Klagenfurt, AT": (46.6247, 14.3055), "Villach, AT": (46.6111, 13.8558), "Wels, AT": (48.1575, 14.0289), "Sankt Polten, AT": (48.2047, 15.6256),
    # BE
    "Brussels, BE": (50.8503, 4.3517), "Antwerp, BE": (51.2194, 4.4025),
    "Ghent, BE": (51.0543, 3.7174), "Charleroi, BE": (50.4108, 4.4446), "Liege, BE": (50.6326, 5.5797), "Bruges, BE": (51.2093, 3.2247), "Namur, BE": (50.4669, 4.8675), "Leuven, BE": (50.8798, 4.7005), "Mons, BE": (50.4542, 3.9563),
    # DK
    "Copenhagen, DK": (55.6761, 12.5683), "Aarhus, DK": (56.1629, 10.2039),
    "Odense, DK": (55.4038, 10.4024), "Aalborg, DK": (57.0488, 9.9217), "Esbjerg, DK": (55.4765, 8.4594), "Randers, DK": (56.4607, 10.0369), "Kolding, DK": (55.4904, 9.4721), "Horsens, DK": (55.8607, 9.8503), "Vejle, DK": (55.7091, 9.5357), "Roskilde, DK": (55.6415, 12.0803),
    # NO
    "Oslo, NO": (59.9139, 10.7522), "Bergen, NO": (60.3913, 5.3221),
    "Trondheim, NO": (63.4305, 10.3951), "Stavanger, NO": (58.97, 5.7331), "Baerum, NO": (59.894, 10.521), "Kristiansand, NO": (58.1467, 7.9956), "Fredrikstad, NO": (59.2181, 10.9298), "Tromso, NO": (69.6492, 18.9553), "Drammen, NO": (59.744, 10.2045), "Skien, NO": (59.2098, 9.6085),
    # FI
    "Helsinki, FI": (60.1699, 24.9384), "Tampere, FI": (61.4978, 23.7610),
    "Turku, FI": (60.4518, 22.2666), "Oulu, FI": (65.0121, 25.4651), "Jyvaskyla, FI": (62.2426, 25.7473), "Lahti, FI": (60.9827, 25.6612), "Kuopio, FI": (62.8924, 27.677), "Pori, FI": (61.4851, 21.7972), "Kouvola, FI": (60.8679, 26.7042), "Vaasa, FI": (63.0951, 21.6165),
    # CZ
    "Prague, CZ": (50.0755, 14.4378), "Brno, CZ": (49.1951, 16.6068),
    "Ostrava, CZ": (49.8209, 18.2625), "Plzen, CZ": (49.7384, 13.3736), "Liberec, CZ": (50.7663, 15.0543), "Olomouc, CZ": (49.5938, 17.2509), "Usti nad Labem, CZ": (50.6607, 14.0328), "Ceske Budejovice, CZ": (48.9747, 14.4744), "Hradec Kralove, CZ": (50.2092, 15.8328), "Pardubice, CZ": (50.0343, 15.7812),
    # RO
    "Bucharest, RO": (44.4268, 26.1025), "Cluj-Napoca, RO": (46.7712, 23.6236),
    "Timisoara, RO": (45.7489, 21.2087), "Iasi, RO": (47.1585, 27.6014), "Constanta, RO": (44.1598, 28.6348), "Craiova, RO": (44.3302, 23.7949), "Brasov, RO": (45.6427, 25.5887), "Galati, RO": (45.4353, 28.008), "Oradea, RO": (47.0722, 21.9217), "Ploiesti, RO": (44.9366, 26.0225),
    # GR
    "Athens, GR": (37.9838, 23.7275), "Thessaloniki, GR": (40.6401, 22.9444),
    "Patras, GR": (38.2466, 21.7346), "Heraklion, GR": (35.3387, 25.1442), "Larissa, GR": (39.639, 22.4191), "Volos, GR": (39.3622, 22.9425), "Ioannina, GR": (39.665, 20.8537), "Chania, GR": (35.5138, 24.018), "Rhodes, GR": (36.4341, 28.2176), "Kavala, GR": (40.9397, 24.4021),
    # HU
    "Budapest, HU": (47.4979, 19.0402),
    "Debrecen, HU": (47.5316, 21.6273), "Szeged, HU": (46.253, 20.1414), "Miskolc, HU": (48.1035, 20.7784), "Pecs, HU": (46.0727, 18.2323), "Gyor, HU": (47.6875, 17.6504), "Nyiregyhaza, HU": (47.9554, 21.7167), "Kecskemet, HU": (46.9062, 19.6913), "Szekesfehervar, HU": (47.186, 18.4221), "Szombathely, HU": (47.2307, 16.6218),
    # IL
    "Tel Aviv, IL": (32.0853, 34.7818), "Jerusalem, IL": (31.7683, 35.2137),
    "Haifa, IL": (32.794, 34.9896), "Rishon LeZion, IL": (31.973, 34.8066), "Petah Tikva, IL": (32.0917, 34.885), "Ashdod, IL": (31.8014, 34.6435), "Netanya, IL": (32.3328, 34.86), "Beersheba, IL": (31.253, 34.7915), "Holon, IL": (32.0158, 34.7874), "Bnei Brak, IL": (32.0807, 34.8338),
    # HK
    "Hong Kong, HK": (22.3193, 114.1694),
    "Kowloon, HK": (22.3193, 114.1694), "Tsuen Wan, HK": (22.372, 114.1137),
    # TW
    "Taipei, TW": (25.0330, 121.5654), "Kaohsiung, TW": (22.6273, 120.3014),
    "Taichung, TW": (24.1477, 120.6736), "Tainan, TW": (22.9998, 120.2269), "New Taipei, TW": (25.0169, 121.4628), "Hsinchu, TW": (24.8138, 120.9675), "Keelung, TW": (25.1276, 121.7391), "Chiayi, TW": (23.4801, 120.4491), "Taoyuan, TW": (24.9936, 121.301), "Changhua, TW": (24.0736, 120.5432),
    # BD
    "Dhaka, BD": (23.8103, 90.4125), "Chittagong, BD": (22.3569, 91.7832),
    "Khulna, BD": (22.8456, 89.5403), "Rajshahi, BD": (24.3745, 88.6042), "Sylhet, BD": (24.8949, 91.8687), "Barisal, BD": (22.701, 90.3535), "Rangpur, BD": (25.7439, 89.2752), "Comilla, BD": (23.4607, 91.1809), "Mymensingh, BD": (24.7471, 90.4203), "Narayanganj, BD": (23.6238, 90.5),
    # LK
    "Colombo, LK": (6.9271, 79.8612),
    "Kandy, LK": (7.2906, 80.6337), "Galle, LK": (6.0535, 80.221), "Jaffna, LK": (9.6615, 80.0255), "Negombo, LK": (7.2083, 79.8358), "Trincomalee, LK": (8.5874, 81.2152), "Anuradhapura, LK": (8.3114, 80.4037), "Batticaloa, LK": (7.717, 81.7), "Matara, LK": (5.9549, 80.555), "Kurunegala, LK": (7.4863, 80.3647),
    # NP
    "Kathmandu, NP": (27.7172, 85.3240),
    "Pokhara, NP": (28.2096, 83.9856), "Lalitpur, NP": (27.6588, 85.3247), "Bharatpur, NP": (27.6833, 84.4333), "Biratnagar, NP": (26.4525, 87.2718), "Birgunj, NP": (27.0104, 84.877), "Dharan, NP": (26.8065, 87.2846), "Butwal, NP": (27.7, 83.4486), "Nepalgunj, NP": (28.05, 81.6167), "Hetauda, NP": (27.4287, 85.0325),
}
# Backward compat alias
CITY_GEO = CITY_COORDS


def google_domain(country: str) -> str:
    # Bare domain (no leading www) - callers prepend "www." themselves.
    return GOOGLE_DOMAINS.get((country or "us").lower(), "google.com")



def encode_uule(city_key: str) -> str | None:
    """Encode a city key into a Google UULE v2 parameter value.
    UULE tells Google's servers which city to use for localised results —
    unlike CDP geolocation which only affects the JS sensor."""
    import base64
    canonical = CITY_CANONICAL.get(city_key)
    if not canonical:
        return None
    encoded = base64.b64encode(canonical.encode("utf-8")).decode()
    return "w+CAIQICI" + encoded


# Canonical city names for Google UULE encoding - single source of truth
# Keys: "CityName, CC" (used in dropdown). Values: Google Ads canonical name.
CITY_CANONICAL = {
    # United States (50 cities)
    "New York, US": "New York,New York,United States",
    "Los Angeles, US": "Los Angeles,California,United States",
    "Chicago, US": "Chicago,Illinois,United States",
    "Houston, US": "Houston,Texas,United States",
    "Miami, US": "Miami,Florida,United States",
    "San Francisco, US": "San Francisco,California,United States",
    "Seattle, US": "Seattle,Washington,United States",
    "Dallas, US": "Dallas,Texas,United States",
    "Denver, US": "Denver,Colorado,United States",
    "Boston, US": "Boston,Massachusetts,United States",
    "Atlanta, US": "Atlanta,Georgia,United States",
    "Phoenix, US": "Phoenix,Arizona,United States",
    "Philadelphia, US": "Philadelphia,Pennsylvania,United States",
    "San Diego, US": "San Diego,California,United States",
    "Austin, US": "Austin,Texas,United States",
    "Las Vegas, US": "Las Vegas,Nevada,United States",
    "Portland, US": "Portland,Oregon,United States",
    "Nashville, US": "Nashville,Tennessee,United States",
    "Charlotte, US": "Charlotte,North Carolina,United States",
    "San Antonio, US": "San Antonio,Texas,United States",
    "Columbus, US": "Columbus,Ohio,United States",
    "Indianapolis, US": "Indianapolis,Indiana,United States",
    "Jacksonville, US": "Jacksonville,Florida,United States",
    "Fort Worth, US": "Fort Worth,Texas,United States",
    "San Jose, US": "San Jose,California,United States",
    "Detroit, US": "Detroit,Michigan,United States",
    "Memphis, US": "Memphis,Tennessee,United States",
    "Baltimore, US": "Baltimore,Maryland,United States",
    "Milwaukee, US": "Milwaukee,Wisconsin,United States",
    "Tucson, US": "Tucson,Arizona,United States",
    "Sacramento, US": "Sacramento,California,United States",
    "Kansas City, US": "Kansas City,Missouri,United States",
    "Raleigh, US": "Raleigh,North Carolina,United States",
    "Tampa, US": "Tampa,Florida,United States",
    "Minneapolis, US": "Minneapolis,Minnesota,United States",
    "Orlando, US": "Orlando,Florida,United States",
    "Pittsburgh, US": "Pittsburgh,Pennsylvania,United States",
    "Cincinnati, US": "Cincinnati,Ohio,United States",
    "New Orleans, US": "New Orleans,Louisiana,United States",
    "Salt Lake City, US": "Salt Lake City,Utah,United States",
    "Honolulu, US": "Honolulu,Hawaii,United States",
    "Anchorage, US": "Anchorage,Alaska,United States",
    "St. Louis, US": "St. Louis,Missouri,United States",
    "Cleveland, US": "Cleveland,Ohio,United States",
    "Virginia Beach, US": "Virginia Beach,Virginia,United States",
    "Omaha, US": "Omaha,Nebraska,United States",
    "Oakland, US": "Oakland,California,United States",
    "Tulsa, US": "Tulsa,Oklahoma,United States",
    "Bakersfield, US": "Bakersfield,California,United States",
    "Albuquerque, US": "Albuquerque,New Mexico,United States",
    # Australia (20 cities)
    "Sydney, AU": "Sydney,New South Wales,Australia",
    "Melbourne, AU": "Melbourne,Victoria,Australia",
    "Brisbane, AU": "Brisbane,Queensland,Australia",
    "Perth, AU": "Perth,Western Australia,Australia",
    "Adelaide, AU": "Adelaide,South Australia,Australia",
    "Gold Coast, AU": "Gold Coast,Queensland,Australia",
    "Canberra, AU": "Canberra,Australian Capital Territory,Australia",
    "Newcastle, AU": "Newcastle,New South Wales,Australia",
    "Hobart, AU": "Hobart,Tasmania,Australia",
    "Darwin, AU": "Darwin,Northern Territory,Australia",
    "Wollongong, AU": "Wollongong,New South Wales,Australia",
    "Cairns, AU": "Cairns,Queensland,Australia",
    "Townsville, AU": "Townsville,Queensland,Australia",
    "Geelong, AU": "Geelong,Victoria,Australia",
    "Sunshine Coast, AU": "Sunshine Coast,Queensland,Australia",
    "Toowoomba, AU": "Toowoomba,Queensland,Australia",
    "Ballarat, AU": "Ballarat,Victoria,Australia",
    "Bendigo, AU": "Bendigo,Victoria,Australia",
    "Launceston, AU": "Launceston,Tasmania,Australia",
    "Rockhampton, AU": "Rockhampton,Queensland,Australia",
    # Canada (20 cities)
    "Toronto, CA": "Toronto,Ontario,Canada",
    "Vancouver, CA": "Vancouver,British Columbia,Canada",
    "Montreal, CA": "Montreal,Quebec,Canada",
    "Calgary, CA": "Calgary,Alberta,Canada",
    "Ottawa, CA": "Ottawa,Ontario,Canada",
    "Edmonton, CA": "Edmonton,Alberta,Canada",
    "Winnipeg, CA": "Winnipeg,Manitoba,Canada",
    "Quebec City, CA": "Quebec City,Quebec,Canada",
    "Hamilton, CA": "Hamilton,Ontario,Canada",
    "Victoria, CA": "Victoria,British Columbia,Canada",
    "Halifax, CA": "Halifax,Nova Scotia,Canada",
    "Saskatoon, CA": "Saskatoon,Saskatchewan,Canada",
    "Regina, CA": "Regina,Saskatchewan,Canada",
    "Mississauga, CA": "Mississauga,Ontario,Canada",
    "Brampton, CA": "Brampton,Ontario,Canada",
    "Kitchener, CA": "Kitchener,Ontario,Canada",
    "London, CA": "London,Ontario,Canada",
    "Surrey, CA": "Surrey,British Columbia,Canada",
    "Markham, CA": "Markham,Ontario,Canada",
    "St. John's, CA": "St. John's,Newfoundland and Labrador,Canada",
    "Fredericton, CA": "Fredericton,New Brunswick,Canada",
    "Charlottetown, CA": "Charlottetown,Prince Edward Island,Canada",
    "Whitehorse, CA": "Whitehorse,Yukon,Canada",
    "Yellowknife, CA": "Yellowknife,Northwest Territories,Canada",
    "Iqaluit, CA": "Iqaluit,Nunavut,Canada",
    # United Kingdom (15 cities)
    "London, GB": "London,England,United Kingdom",
    "Manchester, GB": "Manchester,England,United Kingdom",
    "Birmingham, GB": "Birmingham,England,United Kingdom",
    "Glasgow, GB": "Glasgow,Scotland,United Kingdom",
    "Edinburgh, GB": "Edinburgh,Scotland,United Kingdom",
    "Liverpool, GB": "Liverpool,England,United Kingdom",
    "Leeds, GB": "Leeds,England,United Kingdom",
    "Bristol, GB": "Bristol,England,United Kingdom",
    "Sheffield, GB": "Sheffield,England,United Kingdom",
    "Cardiff, GB": "Cardiff,Wales,United Kingdom",
    "Belfast, GB": "Belfast,Northern Ireland,United Kingdom",
    "Newcastle, GB": "Newcastle upon Tyne,England,United Kingdom",
    "Nottingham, GB": "Nottingham,England,United Kingdom",
    "Southampton, GB": "Southampton,England,United Kingdom",
    "Cambridge, GB": "Cambridge,England,United Kingdom",
    # India (20 cities)
    "Mumbai, IN": "Mumbai,Maharashtra,India",
    "New Delhi, IN": "New Delhi,Delhi,India",
    "Bangalore, IN": "Bangalore,Karnataka,India",
    "Chennai, IN": "Chennai,Tamil Nadu,India",
    "Hyderabad, IN": "Hyderabad,Telangana,India",
    "Kolkata, IN": "Kolkata,West Bengal,India",
    "Pune, IN": "Pune,Maharashtra,India",
    "Jaipur, IN": "Jaipur,Rajasthan,India",
    "Ahmedabad, IN": "Ahmedabad,Gujarat,India",
    "Lucknow, IN": "Lucknow,Uttar Pradesh,India",
    "Chandigarh, IN": "Chandigarh,India",
    "Indore, IN": "Indore,Madhya Pradesh,India",
    "Bhopal, IN": "Bhopal,Madhya Pradesh,India",
    "Nagpur, IN": "Nagpur,Maharashtra,India",
    "Kochi, IN": "Kochi,Kerala,India",
    "Coimbatore, IN": "Coimbatore,Tamil Nadu,India",
    "Surat, IN": "Surat,Gujarat,India",
    "Visakhapatnam, IN": "Visakhapatnam,Andhra Pradesh,India",
    "Patna, IN": "Patna,Bihar,India",
    "Thiruvananthapuram, IN": "Thiruvananthapuram,Kerala,India",
    "Itanagar, IN": "Itanagar,Arunachal Pradesh,India",
    "Guwahati, IN": "Guwahati,Assam,India",
    "Raipur, IN": "Raipur,Chhattisgarh,India",
    "Panaji, IN": "Panaji,Goa,India",
    "Gurugram, IN": "Gurugram,Haryana,India",
    "Shimla, IN": "Shimla,Himachal Pradesh,India",
    "Ranchi, IN": "Ranchi,Jharkhand,India",
    "Imphal, IN": "Imphal,Manipur,India",
    "Shillong, IN": "Shillong,Meghalaya,India",
    "Aizawl, IN": "Aizawl,Mizoram,India",
    "Kohima, IN": "Kohima,Nagaland,India",
    "Bhubaneswar, IN": "Bhubaneswar,Odisha,India",
    "Amritsar, IN": "Amritsar,Punjab,India",
    "Gangtok, IN": "Gangtok,Sikkim,India",
    "Agartala, IN": "Agartala,Tripura,India",
    "Dehradun, IN": "Dehradun,Uttarakhand,India",
    "Port Blair, IN": "Port Blair,Andaman and Nicobar Islands,India",
    "Daman, IN": "Daman,Dadra and Nagar Haveli and Daman and Diu,India",
    "Srinagar, IN": "Srinagar,Jammu and Kashmir,India",
    "Leh, IN": "Leh,Ladakh,India",
    "Kavaratti, IN": "Kavaratti,Lakshadweep,India",
    "Puducherry, IN": "Puducherry,India",
    # Germany (10 cities)
    "Berlin, DE": "Berlin,Germany",
    "Munich, DE": "Munich,Bavaria,Germany",
    "Hamburg, DE": "Hamburg,Germany",
    "Frankfurt, DE": "Frankfurt,Hesse,Germany",
    "Cologne, DE": "Cologne,North Rhine-Westphalia,Germany",
    "Stuttgart, DE": "Stuttgart,Baden-Wurttemberg,Germany",
    "Dusseldorf, DE": "Dusseldorf,North Rhine-Westphalia,Germany",
    "Dortmund, DE": "Dortmund,North Rhine-Westphalia,Germany",
    "Essen, DE": "Essen,North Rhine-Westphalia,Germany",
    "Leipzig, DE": "Leipzig,Saxony,Germany",
    "Potsdam, DE": "Potsdam,Brandenburg,Germany",
    "Bremen, DE": "Bremen,Germany",
    "Hanover, DE": "Hanover,Lower Saxony,Germany",
    "Rostock, DE": "Rostock,Mecklenburg-Vorpommern,Germany",
    "Mainz, DE": "Mainz,Rhineland-Palatinate,Germany",
    "Saarbrucken, DE": "Saarbrucken,Saarland,Germany",
    "Magdeburg, DE": "Magdeburg,Saxony-Anhalt,Germany",
    "Kiel, DE": "Kiel,Schleswig-Holstein,Germany",
    "Erfurt, DE": "Erfurt,Thuringia,Germany",
    # France (8 cities)
    "Paris, FR": "Paris,Ile-de-France,France",
    "Lyon, FR": "Lyon,Auvergne-Rhone-Alpes,France",
    "Marseille, FR": "Marseille,Provence-Alpes-Cote d'Azur,France",
    "Toulouse, FR": "Toulouse,Occitanie,France",
    "Nice, FR": "Nice,Provence-Alpes-Cote d'Azur,France",
    "Bordeaux, FR": "Bordeaux,Nouvelle-Aquitaine,France",
    "Lille, FR": "Lille,Hauts-de-France,France",
    "Strasbourg, FR": "Strasbourg,Grand Est,France",
    "Rouen, FR": "Rouen,Normandie,France",
    "Rennes, FR": "Rennes,Bretagne,France",
    "Nantes, FR": "Nantes,Pays de la Loire,France",
    "Orleans, FR": "Orleans,Centre-Val de Loire,France",
    "Dijon, FR": "Dijon,Bourgogne-Franche-Comte,France",
    "Ajaccio, FR": "Ajaccio,Corse,France",
    # Spain (6 cities)
    "Madrid, ES": "Madrid,Community of Madrid,Spain",
    "Barcelona, ES": "Barcelona,Catalonia,Spain",
    "Valencia, ES": "Valencia,Valencian Community,Spain",
    "Seville, ES": "Seville,Andalusia,Spain",
    "Malaga, ES": "Malaga,Andalusia,Spain",
    "Bilbao, ES": "Bilbao,Basque Country,Spain",
    "Zaragoza, ES": "Zaragoza,Spain",
    "Murcia, ES": "Murcia,Spain",
    "Palma, ES": "Palma,Spain",
    "Las Palmas, ES": "Las Palmas,Spain",
    # Italy (6 cities)
    "Rome, IT": "Rome,Lazio,Italy",
    "Milan, IT": "Milan,Lombardy,Italy",
    "Naples, IT": "Naples,Campania,Italy",
    "Turin, IT": "Turin,Piedmont,Italy",
    "Florence, IT": "Florence,Tuscany,Italy",
    "Bologna, IT": "Bologna,Emilia-Romagna,Italy",
    "Venice, IT": "Venice,Veneto,Italy",
    "Palermo, IT": "Palermo,Sicily,Italy",
    "Bari, IT": "Bari,Puglia,Italy",
    "Catanzaro, IT": "Catanzaro,Calabria,Italy",
    "Cagliari, IT": "Cagliari,Sardinia,Italy",
    "Genoa, IT": "Genoa,Liguria,Italy",
    "Ancona, IT": "Ancona,Marche,Italy",
    "Pescara, IT": "Pescara,Abruzzo,Italy",
    "Trieste, IT": "Trieste,Friuli-Venezia Giulia,Italy",
    "Trento, IT": "Trento,Trentino-Alto Adige,Italy",
    "Perugia, IT": "Perugia,Umbria,Italy",
    "Potenza, IT": "Potenza,Basilicata,Italy",
    "Campobasso, IT": "Campobasso,Molise,Italy",
    "Aosta, IT": "Aosta,Valle d'Aosta,Italy",
    # Netherlands (4 cities)
    "Amsterdam, NL": "Amsterdam,North Holland,Netherlands",
    "Rotterdam, NL": "Rotterdam,South Holland,Netherlands",
    "The Hague, NL": "The Hague,South Holland,Netherlands",
    "Utrecht, NL": "Utrecht,Utrecht,Netherlands",
    "Eindhoven, NL": "Eindhoven,Netherlands",
    "Tilburg, NL": "Tilburg,Netherlands",
    "Groningen, NL": "Groningen,Netherlands",
    "Almere, NL": "Almere,Netherlands",
    "Breda, NL": "Breda,Netherlands",
    "Nijmegen, NL": "Nijmegen,Netherlands",
    # Brazil (5 cities)
    "Sao Paulo, BR": "Sao Paulo,State of Sao Paulo,Brazil",
    "Rio de Janeiro, BR": "Rio de Janeiro,State of Rio de Janeiro,Brazil",
    "Brasilia, BR": "Brasilia,Federal District,Brazil",
    "Salvador, BR": "Salvador,State of Bahia,Brazil",
    "Curitiba, BR": "Curitiba,State of Parana,Brazil",
    "Fortaleza, BR": "Fortaleza,Brazil",
    "Belo Horizonte, BR": "Belo Horizonte,Brazil",
    "Manaus, BR": "Manaus,Brazil",
    "Recife, BR": "Recife,Brazil",
    "Porto Alegre, BR": "Porto Alegre,Brazil",
    # Mexico (5 cities)
    "Mexico City, MX": "Mexico City,Mexico",
    "Guadalajara, MX": "Guadalajara,Jalisco,Mexico",
    "Monterrey, MX": "Monterrey,Nuevo Leon,Mexico",
    "Cancun, MX": "Cancun,Quintana Roo,Mexico",
    "Puebla, MX": "Puebla,Puebla,Mexico",
    "Tijuana, MX": "Tijuana,Mexico",
    "Leon, MX": "Leon,Mexico",
    "Juarez, MX": "Juarez,Mexico",
    "Zapopan, MX": "Zapopan,Mexico",
    "Merida, MX": "Merida,Mexico",
    # Japan (5 cities)
    "Tokyo, JP": "Tokyo,Japan",
    "Osaka, JP": "Osaka,Osaka,Japan",
    "Yokohama, JP": "Yokohama,Kanagawa,Japan",
    "Nagoya, JP": "Nagoya,Aichi,Japan",
    "Kyoto, JP": "Kyoto,Kyoto,Japan",
    "Sapporo, JP": "Sapporo,Japan",
    "Fukuoka, JP": "Fukuoka,Japan",
    "Kobe, JP": "Kobe,Japan",
    "Kawasaki, JP": "Kawasaki,Japan",
    "Saitama, JP": "Saitama,Japan",
    # South Korea (4 cities)
    "Seoul, KR": "Seoul,South Korea",
    "Busan, KR": "Busan,South Korea",
    "Incheon, KR": "Incheon,South Korea",
    "Daegu, KR": "Daegu,South Korea",
    "Daejeon, KR": "Daejeon,South Korea",
    "Gwangju, KR": "Gwangju,South Korea",
    "Suwon, KR": "Suwon,South Korea",
    "Ulsan, KR": "Ulsan,South Korea",
    "Changwon, KR": "Changwon,South Korea",
    "Goyang, KR": "Goyang,South Korea",
    # Russia (3 cities)
    "Moscow, RU": "Moscow,Russia",
    "Saint Petersburg, RU": "Saint Petersburg,Russia",
    "Novosibirsk, RU": "Novosibirsk,Novosibirsk Oblast,Russia",
    "Yekaterinburg, RU": "Yekaterinburg,Russia",
    "Kazan, RU": "Kazan,Russia",
    "Nizhny Novgorod, RU": "Nizhny Novgorod,Russia",
    "Chelyabinsk, RU": "Chelyabinsk,Russia",
    "Samara, RU": "Samara,Russia",
    "Omsk, RU": "Omsk,Russia",
    "Rostov-on-Don, RU": "Rostov-on-Don,Russia",
    # UAE (3 cities)
    "Dubai, AE": "Dubai,United Arab Emirates",
    "Abu Dhabi, AE": "Abu Dhabi,United Arab Emirates",
    "Sharjah, AE": "Sharjah,United Arab Emirates",
    # Singapore
    "Singapore, SG": "Singapore",
    "Jurong, SG": "Jurong,Singapore",
    "Tampines, SG": "Tampines,Singapore",
    "Woodlands, SG": "Woodlands,Singapore",
    "Ajman, AE": "Ajman,United Arab Emirates",
    "Ras Al Khaimah, AE": "Ras Al Khaimah,United Arab Emirates",
    "Fujairah, AE": "Fujairah,United Arab Emirates",
    "Umm Al Quwain, AE": "Umm Al Quwain,United Arab Emirates",
    "Al Ain, AE": "Al Ain,United Arab Emirates",
    # South Africa (3 cities)
    "Johannesburg, ZA": "Johannesburg,Gauteng,South Africa",
    "Cape Town, ZA": "Cape Town,Western Cape,South Africa",
    "Durban, ZA": "Durban,KwaZulu-Natal,South Africa",
    "Pretoria, ZA": "Pretoria,South Africa",
    "Port Elizabeth, ZA": "Port Elizabeth,South Africa",
    "Bloemfontein, ZA": "Bloemfontein,South Africa",
    "East London, ZA": "East London,South Africa",
    "Nelspruit, ZA": "Nelspruit,South Africa",
    "Polokwane, ZA": "Polokwane,South Africa",
    "Kimberley, ZA": "Kimberley,South Africa",
    # New Zealand (4 cities)
    "Auckland, NZ": "Auckland,New Zealand",
    "Wellington, NZ": "Wellington,New Zealand",
    "Christchurch, NZ": "Christchurch,Canterbury,New Zealand",
    "Hamilton, NZ": "Hamilton,Waikato,New Zealand",
    "Tauranga, NZ": "Tauranga,New Zealand",
    "Napier, NZ": "Napier,New Zealand",
    "Dunedin, NZ": "Dunedin,New Zealand",
    "Palmerston North, NZ": "Palmerston North,New Zealand",
    "Nelson, NZ": "Nelson,New Zealand",
    "Rotorua, NZ": "Rotorua,New Zealand",
    # Ireland (3 cities)
    "Dublin, IE": "Dublin,Ireland",
    "Cork, IE": "Cork,Ireland",
    "Galway, IE": "Galway,Ireland",
    "Limerick, IE": "Limerick,Ireland",
    "Waterford, IE": "Waterford,Ireland",
    "Drogheda, IE": "Drogheda,Ireland",
    "Kilkenny, IE": "Kilkenny,Ireland",
    "Wexford, IE": "Wexford,Ireland",
    # Sweden (3 cities)
    "Stockholm, SE": "Stockholm,Sweden",
    "Gothenburg, SE": "Gothenburg,Sweden",
    "Malmo, SE": "Malmo,Sweden",
    "Uppsala, SE": "Uppsala,Sweden",
    "Vasteras, SE": "Vasteras,Sweden",
    "Orebro, SE": "Orebro,Sweden",
    "Linkoping, SE": "Linkoping,Sweden",
    "Helsingborg, SE": "Helsingborg,Sweden",
    "Jonkoping, SE": "Jonkoping,Sweden",
    "Norrkoping, SE": "Norrkoping,Sweden",
    # Switzerland (3 cities)
    "Zurich, CH": "Zurich,Switzerland",
    "Geneva, CH": "Geneva,Switzerland",
    "Bern, CH": "Bern,Switzerland",
    "Basel, CH": "Basel,Switzerland",
    "Lausanne, CH": "Lausanne,Switzerland",
    "Lucerne, CH": "Lucerne,Switzerland",
    "Winterthur, CH": "Winterthur,Switzerland",
    "St. Gallen, CH": "St. Gallen,Switzerland",
    "Lugano, CH": "Lugano,Switzerland",
    "Biel, CH": "Biel,Switzerland",
    # Poland (3 cities)
    "Warsaw, PL": "Warsaw,Masovian Voivodeship,Poland",
    "Krakow, PL": "Krakow,Lesser Poland Voivodeship,Poland",
    "Wroclaw, PL": "Wroclaw,Lower Silesian Voivodeship,Poland",
    "Lodz, PL": "Lodz,Poland",
    "Poznan, PL": "Poznan,Poland",
    "Gdansk, PL": "Gdansk,Poland",
    "Szczecin, PL": "Szczecin,Poland",
    "Bydgoszcz, PL": "Bydgoszcz,Poland",
    "Lublin, PL": "Lublin,Poland",
    "Katowice, PL": "Katowice,Poland",
    # Indonesia (4 cities)
    "Jakarta, ID": "Jakarta,Indonesia",
    "Surabaya, ID": "Surabaya,East Java,Indonesia",
    "Bali, ID": "Bali,Indonesia",
    "Bandung, ID": "Bandung,West Java,Indonesia",
    "Medan, ID": "Medan,Indonesia",
    "Semarang, ID": "Semarang,Indonesia",
    "Makassar, ID": "Makassar,Indonesia",
    "Palembang, ID": "Palembang,Indonesia",
    "Yogyakarta, ID": "Yogyakarta,Indonesia",
    "Denpasar, ID": "Denpasar,Indonesia",
    # Philippines (3 cities)
    "Manila, PH": "Manila,Metro Manila,Philippines",
    "Cebu, PH": "Cebu City,Central Visayas,Philippines",
    "Davao, PH": "Davao City,Davao Region,Philippines",
    "Quezon City, PH": "Quezon City,Philippines",
    "Zamboanga, PH": "Zamboanga,Philippines",
    "Antipolo, PH": "Antipolo,Philippines",
    "Taguig, PH": "Taguig,Philippines",
    "Cagayan de Oro, PH": "Cagayan de Oro,Philippines",
    "Paranaque, PH": "Paranaque,Philippines",
    "Makati, PH": "Makati,Philippines",
    # Malaysia (3 cities)
    "Kuala Lumpur, MY": "Kuala Lumpur,Federal Territory of Kuala Lumpur,Malaysia",
    "Penang, MY": "George Town,Penang,Malaysia",
    "Johor Bahru, MY": "Johor Bahru,Johor,Malaysia",
    "Ipoh, MY": "Ipoh,Malaysia",
    "Shah Alam, MY": "Shah Alam,Malaysia",
    "Kuching, MY": "Kuching,Malaysia",
    "Kota Kinabalu, MY": "Kota Kinabalu,Malaysia",
    "Malacca City, MY": "Malacca City,Malaysia",
    "Alor Setar, MY": "Alor Setar,Malaysia",
    "Petaling Jaya, MY": "Petaling Jaya,Malaysia",
    # Pakistan (4 cities)
    "Karachi, PK": "Karachi,Sindh,Pakistan",
    "Lahore, PK": "Lahore,Punjab,Pakistan",
    "Islamabad, PK": "Islamabad,Islamabad Capital Territory,Pakistan",
    "Rawalpindi, PK": "Rawalpindi,Punjab,Pakistan",
    "Faisalabad, PK": "Faisalabad,Pakistan",
    "Multan, PK": "Multan,Pakistan",
    "Peshawar, PK": "Peshawar,Pakistan",
    "Quetta, PK": "Quetta,Pakistan",
    "Sialkot, PK": "Sialkot,Pakistan",
    "Gujranwala, PK": "Gujranwala,Pakistan",
    # Nigeria (3 cities)
    "Lagos, NG": "Lagos,Nigeria",
    "Abuja, NG": "Abuja,Nigeria",
    "Port Harcourt, NG": "Port Harcourt,Rivers,Nigeria",
    "Kano, NG": "Kano,Nigeria",
    "Ibadan, NG": "Ibadan,Nigeria",
    "Benin City, NG": "Benin City,Nigeria",
    "Kaduna, NG": "Kaduna,Nigeria",
    "Enugu, NG": "Enugu,Nigeria",
    "Zaria, NG": "Zaria,Nigeria",
    "Aba, NG": "Aba,Nigeria",
    # Saudi Arabia (4 cities)
    "Riyadh, SA": "Riyadh,Riyadh Province,Saudi Arabia",
    "Jeddah, SA": "Jeddah,Makkah Province,Saudi Arabia",
    "Mecca, SA": "Mecca,Makkah Province,Saudi Arabia",
    "Dammam, SA": "Dammam,Eastern Province,Saudi Arabia",
    "Medina, SA": "Medina,Saudi Arabia",
    "Taif, SA": "Taif,Saudi Arabia",
    "Tabuk, SA": "Tabuk,Saudi Arabia",
    "Buraidah, SA": "Buraidah,Saudi Arabia",
    "Khobar, SA": "Khobar,Saudi Arabia",
    "Abha, SA": "Abha,Saudi Arabia",
    # Turkey (4 cities)
    "Istanbul, TR": "Istanbul,Turkey",
    "Ankara, TR": "Ankara,Turkey",
    "Izmir, TR": "Izmir,Turkey",
    "Antalya, TR": "Antalya,Turkey",
    "Bursa, TR": "Bursa,Turkey",
    "Adana, TR": "Adana,Turkey",
    "Gaziantep, TR": "Gaziantep,Turkey",
    "Konya, TR": "Konya,Turkey",
    "Mersin, TR": "Mersin,Turkey",
    "Kayseri, TR": "Kayseri,Turkey",
    # Thailand (3 cities)
    "Bangkok, TH": "Bangkok,Thailand",
    "Chiang Mai, TH": "Chiang Mai,Thailand",
    "Phuket, TH": "Phuket,Thailand",
    "Pattaya, TH": "Pattaya,Thailand",
    "Nonthaburi, TH": "Nonthaburi,Thailand",
    "Nakhon Ratchasima, TH": "Nakhon Ratchasima,Thailand",
    "Khon Kaen, TH": "Khon Kaen,Thailand",
    "Udon Thani, TH": "Udon Thani,Thailand",
    "Hat Yai, TH": "Hat Yai,Thailand",
    "Rayong, TH": "Rayong,Thailand",
    # Vietnam (3 cities)
    "Ho Chi Minh City, VN": "Ho Chi Minh City,Vietnam",
    "Hanoi, VN": "Hanoi,Vietnam",
    "Da Nang, VN": "Da Nang,Vietnam",
    "Can Tho, VN": "Can Tho,Vietnam",
    "Hai Phong, VN": "Hai Phong,Vietnam",
    "Bien Hoa, VN": "Bien Hoa,Vietnam",
    "Nha Trang, VN": "Nha Trang,Vietnam",
    "Hue, VN": "Hue,Vietnam",
    "Vung Tau, VN": "Vung Tau,Vietnam",
    "Quy Nhon, VN": "Quy Nhon,Vietnam",
    # Argentina (3 cities)
    "Buenos Aires, AR": "Buenos Aires,Argentina",
    "Cordoba, AR": "Cordoba,Cordoba,Argentina",
    "Rosario, AR": "Rosario,Santa Fe,Argentina",
    "Mendoza, AR": "Mendoza,Argentina",
    "La Plata, AR": "La Plata,Argentina",
    "San Miguel de Tucuman, AR": "San Miguel de Tucuman,Argentina",
    "Mar del Plata, AR": "Mar del Plata,Argentina",
    "Salta, AR": "Salta,Argentina",
    "Santa Fe, AR": "Santa Fe,Argentina",
    "Neuquen, AR": "Neuquen,Argentina",
    # Colombia (3 cities)
    "Bogota, CO": "Bogota,Colombia",
    "Medellin, CO": "Medellin,Antioquia,Colombia",
    "Cali, CO": "Cali,Valle del Cauca,Colombia",
    "Barranquilla, CO": "Barranquilla,Colombia",
    "Cartagena, CO": "Cartagena,Colombia",
    "Cucuta, CO": "Cucuta,Colombia",
    "Bucaramanga, CO": "Bucaramanga,Colombia",
    "Pereira, CO": "Pereira,Colombia",
    "Santa Marta, CO": "Santa Marta,Colombia",
    "Ibague, CO": "Ibague,Colombia",
    # Chile (2 cities)
    "Santiago, CL": "Santiago,Santiago Metropolitan,Chile",
    "Valparaiso, CL": "Valparaiso,Valparaiso,Chile",
    "Concepcion, CL": "Concepcion,Chile",
    "La Serena, CL": "La Serena,Chile",
    "Antofagasta, CL": "Antofagasta,Chile",
    "Temuco, CL": "Temuco,Chile",
    "Rancagua, CL": "Rancagua,Chile",
    "Talca, CL": "Talca,Chile",
    "Iquique, CL": "Iquique,Chile",
    "Puerto Montt, CL": "Puerto Montt,Chile",
    # Peru (2 cities)
    "Lima, PE": "Lima,Peru",
    "Arequipa, PE": "Arequipa,Peru",
    "Trujillo, PE": "Trujillo,Peru",
    "Chiclayo, PE": "Chiclayo,Peru",
    "Piura, PE": "Piura,Peru",
    "Iquitos, PE": "Iquitos,Peru",
    "Cusco, PE": "Cusco,Peru",
    "Chimbote, PE": "Chimbote,Peru",
    "Huancayo, PE": "Huancayo,Peru",
    "Tacna, PE": "Tacna,Peru",
    # Egypt (2 cities)
    "Cairo, EG": "Cairo,Egypt",
    "Alexandria, EG": "Alexandria,Egypt",
    "Giza, EG": "Giza,Egypt",
    "Shubra El Kheima, EG": "Shubra El Kheima,Egypt",
    "Port Said, EG": "Port Said,Egypt",
    "Suez, EG": "Suez,Egypt",
    "Luxor, EG": "Luxor,Egypt",
    "Mansoura, EG": "Mansoura,Egypt",
    "Tanta, EG": "Tanta,Egypt",
    "Aswan, EG": "Aswan,Egypt",
    # Kenya (2 cities)
    "Nairobi, KE": "Nairobi,Kenya",
    "Mombasa, KE": "Mombasa,Kenya",
    "Kisumu, KE": "Kisumu,Kenya",
    "Nakuru, KE": "Nakuru,Kenya",
    "Eldoret, KE": "Eldoret,Kenya",
    "Thika, KE": "Thika,Kenya",
    "Malindi, KE": "Malindi,Kenya",
    # Portugal (2 cities)
    "Lisbon, PT": "Lisbon,Portugal",
    "Porto, PT": "Porto,Portugal",
    "Vila Nova de Gaia, PT": "Vila Nova de Gaia,Portugal",
    "Amadora, PT": "Amadora,Portugal",
    "Braga, PT": "Braga,Portugal",
    "Coimbra, PT": "Coimbra,Portugal",
    "Funchal, PT": "Funchal,Portugal",
    "Setubal, PT": "Setubal,Portugal",
    "Almada, PT": "Almada,Portugal",
    "Aveiro, PT": "Aveiro,Portugal",
    # Austria (2 cities)
    "Vienna, AT": "Vienna,Austria",
    "Salzburg, AT": "Salzburg,Austria",
    "Graz, AT": "Graz,Austria",
    "Linz, AT": "Linz,Austria",
    "Innsbruck, AT": "Innsbruck,Austria",
    "Klagenfurt, AT": "Klagenfurt,Austria",
    "Villach, AT": "Villach,Austria",
    "Wels, AT": "Wels,Austria",
    "Sankt Polten, AT": "Sankt Polten,Austria",
    # Belgium (2 cities)
    "Brussels, BE": "Brussels,Belgium",
    "Antwerp, BE": "Antwerp,Flanders,Belgium",
    "Ghent, BE": "Ghent,Belgium",
    "Charleroi, BE": "Charleroi,Belgium",
    "Liege, BE": "Liege,Belgium",
    "Bruges, BE": "Bruges,Belgium",
    "Namur, BE": "Namur,Belgium",
    "Leuven, BE": "Leuven,Belgium",
    "Mons, BE": "Mons,Belgium",
    # Denmark (2 cities)
    "Copenhagen, DK": "Copenhagen,Denmark",
    "Aarhus, DK": "Aarhus,Denmark",
    "Odense, DK": "Odense,Denmark",
    "Aalborg, DK": "Aalborg,Denmark",
    "Esbjerg, DK": "Esbjerg,Denmark",
    "Randers, DK": "Randers,Denmark",
    "Kolding, DK": "Kolding,Denmark",
    "Horsens, DK": "Horsens,Denmark",
    "Vejle, DK": "Vejle,Denmark",
    "Roskilde, DK": "Roskilde,Denmark",
    # Norway (2 cities)
    "Oslo, NO": "Oslo,Norway",
    "Bergen, NO": "Bergen,Norway",
    "Trondheim, NO": "Trondheim,Norway",
    "Stavanger, NO": "Stavanger,Norway",
    "Baerum, NO": "Baerum,Norway",
    "Kristiansand, NO": "Kristiansand,Norway",
    "Fredrikstad, NO": "Fredrikstad,Norway",
    "Tromso, NO": "Tromso,Norway",
    "Drammen, NO": "Drammen,Norway",
    "Skien, NO": "Skien,Norway",
    # Finland (2 cities)
    "Helsinki, FI": "Helsinki,Finland",
    "Tampere, FI": "Tampere,Finland",
    "Turku, FI": "Turku,Finland",
    "Oulu, FI": "Oulu,Finland",
    "Jyvaskyla, FI": "Jyvaskyla,Finland",
    "Lahti, FI": "Lahti,Finland",
    "Kuopio, FI": "Kuopio,Finland",
    "Pori, FI": "Pori,Finland",
    "Kouvola, FI": "Kouvola,Finland",
    "Vaasa, FI": "Vaasa,Finland",
    # Czech Republic (2 cities)
    "Prague, CZ": "Prague,Czech Republic",
    "Brno, CZ": "Brno,South Moravian,Czech Republic",
    "Ostrava, CZ": "Ostrava,Czech Republic",
    "Plzen, CZ": "Plzen,Czech Republic",
    "Liberec, CZ": "Liberec,Czech Republic",
    "Olomouc, CZ": "Olomouc,Czech Republic",
    "Usti nad Labem, CZ": "Usti nad Labem,Czech Republic",
    "Ceske Budejovice, CZ": "Ceske Budejovice,Czech Republic",
    "Hradec Kralove, CZ": "Hradec Kralove,Czech Republic",
    "Pardubice, CZ": "Pardubice,Czech Republic",
    # Romania (2 cities)
    "Bucharest, RO": "Bucharest,Romania",
    "Cluj-Napoca, RO": "Cluj-Napoca,Cluj,Romania",
    "Timisoara, RO": "Timisoara,Romania",
    "Iasi, RO": "Iasi,Romania",
    "Constanta, RO": "Constanta,Romania",
    "Craiova, RO": "Craiova,Romania",
    "Brasov, RO": "Brasov,Romania",
    "Galati, RO": "Galati,Romania",
    "Oradea, RO": "Oradea,Romania",
    "Ploiesti, RO": "Ploiesti,Romania",
    # Greece (2 cities)
    "Athens, GR": "Athens,Greece",
    "Thessaloniki, GR": "Thessaloniki,Greece",
    "Patras, GR": "Patras,Greece",
    "Heraklion, GR": "Heraklion,Greece",
    "Larissa, GR": "Larissa,Greece",
    "Volos, GR": "Volos,Greece",
    "Ioannina, GR": "Ioannina,Greece",
    "Chania, GR": "Chania,Greece",
    "Rhodes, GR": "Rhodes,Greece",
    "Kavala, GR": "Kavala,Greece",
    # Hungary (1 city)
    "Budapest, HU": "Budapest,Hungary",
    "Debrecen, HU": "Debrecen,Hungary",
    "Szeged, HU": "Szeged,Hungary",
    "Miskolc, HU": "Miskolc,Hungary",
    "Pecs, HU": "Pecs,Hungary",
    "Gyor, HU": "Gyor,Hungary",
    "Nyiregyhaza, HU": "Nyiregyhaza,Hungary",
    "Kecskemet, HU": "Kecskemet,Hungary",
    "Szekesfehervar, HU": "Szekesfehervar,Hungary",
    "Szombathely, HU": "Szombathely,Hungary",
    # Israel (2 cities)
    "Tel Aviv, IL": "Tel Aviv,Israel",
    "Jerusalem, IL": "Jerusalem,Israel",
    # Hong Kong
    "Hong Kong, HK": "Hong Kong",
    "Kowloon, HK": "Kowloon,Hong Kong",
    "Tsuen Wan, HK": "Tsuen Wan,Hong Kong",
    "Haifa, IL": "Haifa,Israel",
    "Rishon LeZion, IL": "Rishon LeZion,Israel",
    "Petah Tikva, IL": "Petah Tikva,Israel",
    "Ashdod, IL": "Ashdod,Israel",
    "Netanya, IL": "Netanya,Israel",
    "Beersheba, IL": "Beersheba,Israel",
    "Holon, IL": "Holon,Israel",
    "Bnei Brak, IL": "Bnei Brak,Israel",
    # Taiwan (2 cities)
    "Taipei, TW": "Taipei,Taiwan",
    "Kaohsiung, TW": "Kaohsiung,Taiwan",
    "Taichung, TW": "Taichung,Taiwan",
    "Tainan, TW": "Tainan,Taiwan",
    "New Taipei, TW": "New Taipei,Taiwan",
    "Hsinchu, TW": "Hsinchu,Taiwan",
    "Keelung, TW": "Keelung,Taiwan",
    "Chiayi, TW": "Chiayi,Taiwan",
    "Taoyuan, TW": "Taoyuan,Taiwan",
    "Changhua, TW": "Changhua,Taiwan",
    # Bangladesh (2 cities)
    "Dhaka, BD": "Dhaka,Bangladesh",
    "Chittagong, BD": "Chittagong,Bangladesh",
    "Khulna, BD": "Khulna,Bangladesh",
    "Rajshahi, BD": "Rajshahi,Bangladesh",
    "Sylhet, BD": "Sylhet,Bangladesh",
    "Barisal, BD": "Barisal,Bangladesh",
    "Rangpur, BD": "Rangpur,Bangladesh",
    "Comilla, BD": "Comilla,Bangladesh",
    "Mymensingh, BD": "Mymensingh,Bangladesh",
    "Narayanganj, BD": "Narayanganj,Bangladesh",
    # Sri Lanka (1 city)
    "Colombo, LK": "Colombo,Sri Lanka",
    "Kandy, LK": "Kandy,Sri Lanka",
    "Galle, LK": "Galle,Sri Lanka",
    "Jaffna, LK": "Jaffna,Sri Lanka",
    "Negombo, LK": "Negombo,Sri Lanka",
    "Trincomalee, LK": "Trincomalee,Sri Lanka",
    "Anuradhapura, LK": "Anuradhapura,Sri Lanka",
    "Batticaloa, LK": "Batticaloa,Sri Lanka",
    "Matara, LK": "Matara,Sri Lanka",
    "Kurunegala, LK": "Kurunegala,Sri Lanka",
    # Nepal (1 city)
    "Kathmandu, NP": "Kathmandu,Nepal",
    "Pokhara, NP": "Pokhara,Nepal",
    "Lalitpur, NP": "Lalitpur,Nepal",
    "Bharatpur, NP": "Bharatpur,Nepal",
    "Biratnagar, NP": "Biratnagar,Nepal",
    "Birgunj, NP": "Birgunj,Nepal",
    "Dharan, NP": "Dharan,Nepal",
    "Butwal, NP": "Butwal,Nepal",
    "Nepalgunj, NP": "Nepalgunj,Nepal",
    "Hetauda, NP": "Hetauda,Nepal",
}




def set_geolocation(driver, country=None, latitude=None, longitude=None, accuracy=100):
    """Override browser geolocation via CDP. Uses country defaults if no lat/long given."""
    if latitude is not None and longitude is not None:
        lat, lng = float(latitude), float(longitude)
    elif country and country.lower() in COUNTRY_GEO:
        lat, lng = COUNTRY_GEO[country.lower()]
    else:
        return False
    try:
        driver.execute_cdp_cmd("Emulation.setGeolocationOverride", {
            "latitude": lat, "longitude": lng, "accuracy": accuracy
        })
        return True
    except Exception:
        return False


def country_timezone(country: str) -> str:
    return COUNTRY_TZ.get((country or "us").lower(), "America/New_York")


class BrowserClosedError(Exception):
    """Raised when the browser window is closed by the user or crashes."""


class BlockedError(Exception):
    """Raised when Google blocks the session and recovery is needed."""


# --------------------------------------------------------------------------- #
# Proxy pool with rotation
# --------------------------------------------------------------------------- #
class ProxyPool:
    """
    Rotates through a list of proxies. Each proxy is a dict:
        {"type": "http|socks5", "host": "1.2.3.4", "port": "8080",
         "user": "", "pass": ""}
    A single inline proxy (legacy) is treated as a pool of one.
    """

    def __init__(self, proxies=None):
        self.proxies = [p for p in (proxies or []) if p and p.get("host") and p.get("port")]
        self._i = -1

    def __bool__(self):
        return len(self.proxies) > 0

    def next(self):
        if not self.proxies:
            return None
        self._i = (self._i + 1) % len(self.proxies)
        return self.proxies[self._i]

    def current(self):
        if not self.proxies or self._i < 0:
            return None
        return self.proxies[self._i]


# --------------------------------------------------------------------------- #
# Proxy-auth extension (Chrome can't take user:pass on the CLI)
# --------------------------------------------------------------------------- #
def _build_proxy_auth_extension(proxy) -> str | None:
    """Create a tiny MV2 extension that answers the proxy auth challenge.
    Returns the extension dir path, or None if no auth needed."""
    if not (proxy and proxy.get("user") and proxy.get("pass")):
        return None
    scheme = proxy.get("type", "http")
    host = proxy["host"]
    port = int(proxy["port"])
    user = proxy["user"]
    pwd = proxy["pass"]

    manifest = {
        "version": "1.0.0",
        "manifest_version": 2,
        "name": "GRC Proxy Auth",
        "permissions": ["proxy", "tabs", "unlimitedStorage", "storage",
                         "<all_urls>", "webRequest", "webRequestBlocking"],
        "background": {"scripts": ["background.js"]},
        "minimum_chrome_version": "76.0.0",
    }
    background = """
var config = {
  mode: "fixed_servers",
  rules: { singleProxy: { scheme: "%s", host: "%s", port: %d }, bypassList: ["localhost"] }
};
chrome.proxy.settings.set({value: config, scope: "regular"}, function() {});
chrome.webRequest.onAuthRequired.addListener(
  function(details) {
    return { authCredentials: { username: "%s", password: "%s" } };
  },
  { urls: ["<all_urls>"] }, ["blocking"]
);
""" % (scheme, host, port, user, pwd)

    d = tempfile.mkdtemp(prefix="grc_proxy_")
    path = os.path.join(d, "proxy_auth.zip")
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr("manifest.json", json.dumps(manifest))
        zf.writestr("background.js", background)
    return path


# --------------------------------------------------------------------------- #
# Realistic user-agent / viewport pools
# --------------------------------------------------------------------------- #
_VIEWPORTS = [(1366, 768), (1440, 900), (1536, 864), (1600, 900), (1920, 1080)]

_STEALTH_JS = r"""
(function() {
  // --- webdriver flag ---
  Object.defineProperty(navigator, 'webdriver', {get: () => undefined, configurable: true});

  // --- realistic navigator properties ---
  // Don't override navigator.languages - Chrome sets it correctly from --lang flag
  Object.defineProperty(navigator, 'platform',  {get: () => 'Win32'});
  Object.defineProperty(navigator, 'maxTouchPoints', {get: () => 0});
  Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 8});
  Object.defineProperty(navigator, 'deviceMemory', {get: () => 8});

  // --- realistic plugins (Chrome on Windows typically shows these) ---
  const _plugins = [
    {name:'Chrome PDF Plugin',    filename:'internal-pdf-viewer', description:'Portable Document Format', length:1},
    {name:'Chrome PDF Viewer',    filename:'mhjfbmdgcfjbbpaeojofohoefgiehjai', description:'', length:1},
    {name:'Native Client',        filename:'internal-nacl-plugin', description:'', length:2},
  ];
  Object.defineProperty(navigator, 'plugins', {
    get: () => Object.assign(_plugins, {length: _plugins.length, item: i => _plugins[i], namedItem: n => _plugins.find(p => p.name===n) || null}),
    configurable: true,
  });
  Object.defineProperty(navigator, 'mimeTypes', {get: () => ({length:2, item:()=>null, namedItem:()=>null}), configurable: true});

  // --- chrome object (must look exactly like real Chrome) ---
  if (!window.chrome || !window.chrome.runtime) {
    window.chrome = {
      app: {isInstalled: false, InstallState:{DISABLED:'disabled',INSTALLED:'installed',NOT_INSTALLED:'not_installed'}, RunningState:{CANNOT_RUN:'cannot_run',READY_TO_RUN:'ready_to_run',RUNNING:'running'}},
      runtime: {OnInstalledReason:{CHROME_UPDATE:'chrome_update',INSTALL:'install',SHARED_MODULE_UPDATE:'shared_module_update',UPDATE:'update'}, PlatformArch:{ARM:'arm',ARM64:'arm64',MIPS:'mips',MIPS64:'mips64',X86_32:'x86-32',X86_64:'x86-64'}, PlatformNaclArch:{ARM:'arm',MIPS:'mips',MIPS64:'mips64',X86_32:'x86-32',X86_64:'x86-64'}, PlatformOs:{ANDROID:'android',CROS:'cros',LINUX:'linux',MAC:'mac',OPENBSD:'openbsd',WIN:'win'}, RequestUpdateCheckStatus:{NO_UPDATE:'no_update',THROTTLED:'throttled',UPDATE_AVAILABLE:'update_available'}},
      csi: function(){return {startE:Date.now(),onloadT:Date.now(),pageT:Math.random()*500+500,tran:15};},
      loadTimes: function(){return {commitLoadTime:Date.now()/1000,connectionInfo:'h2',finishDocumentLoadTime:0,finishLoadTime:0,firstPaintAfterLoadTime:0,firstPaintTime:0,navigationType:'Other',npnNegotiatedProtocol:'h2',requestTime:Date.now()/1000,startLoadTime:Date.now()/1000,wasAlternateProtocolAvailable:false,wasFetchedViaSpdy:true,wasNpnNegotiated:true};}
    };
  }

  // --- permissions (avoid notification state leak) ---
  const _origQuery = window.navigator.permissions && window.navigator.permissions.query.bind(window.navigator.permissions);
  if (_origQuery) {
    window.navigator.permissions.query = (p) =>
      p && p.name === 'notifications'
        ? Promise.resolve({state: Notification.permission, onchange: null})
        : _origQuery(p);
  }

  // --- WebGL vendor spoof (GPU fingerprint) ---
  try {
    const _gp = WebGLRenderingContext.prototype.getParameter;
    WebGLRenderingContext.prototype.getParameter = function(p) {
      if (p === 37445) return 'Google Inc. (Intel)';
      if (p === 37446) return 'ANGLE (Intel, Intel(R) UHD Graphics Direct3D11 vs_5_0 ps_5_0, D3D11)';
      return _gp.call(this, p);
    };
    const _gp2 = WebGL2RenderingContext.prototype.getParameter;
    WebGL2RenderingContext.prototype.getParameter = function(p) {
      if (p === 37445) return 'Google Inc. (Intel)';
      if (p === 37446) return 'ANGLE (Intel, Intel(R) UHD Graphics Direct3D11 vs_5_0 ps_5_0, D3D11)';
      return _gp2.call(this, p);
    };
  } catch(e) {}

  // --- Canvas noise (unique per session, prevents canvas fingerprinting) ---
  try {
    const _noise = (Math.random() * 0.04) - 0.02;
    const _origToDataURL = HTMLCanvasElement.prototype.toDataURL;
    HTMLCanvasElement.prototype.toDataURL = function(type, quality) {
      const ctx = this.getContext('2d');
      if (ctx) {
        const imageData = ctx.getImageData(0, 0, this.width || 1, this.height || 1);
        for (let i = 0; i < imageData.data.length; i += 4) {
          imageData.data[i]   = Math.min(255, Math.max(0, imageData.data[i]   + Math.round(_noise * 3)));
          imageData.data[i+1] = Math.min(255, Math.max(0, imageData.data[i+1] + Math.round(_noise * 2)));
        }
        ctx.putImageData(imageData, 0, 0);
      }
      return _origToDataURL.call(this, type, quality);
    };
  } catch(e) {}

  // --- AudioContext noise (audio fingerprint) ---
  try {
    const _origGetChannelData = AudioBuffer.prototype.getChannelData;
    AudioBuffer.prototype.getChannelData = function(ch) {
      const data = _origGetChannelData.call(this, ch);
      for (let i = 0; i < data.length; i += 100) {
        data[i] += (Math.random() - 0.5) * 0.0001;
      }
      return data;
    };
  } catch(e) {}
})();
"""


def get_chrome_major_version():
    """Read installed Chrome major version from the Windows registry."""
    try:
        import subprocess
        keys = [
            r"HKCU\SOFTWARE\Google\Chrome\BLBeacon",
            r"HKLM\SOFTWARE\Google\Chrome\BLBeacon",
            r"HKLM\SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon",
        ]
        for key in keys:
            r = subprocess.run(["reg", "query", key, "/v", "version"],
                               capture_output=True, text=True)
            if r.returncode == 0:
                m = re.search(r"(\d+)\.\d+\.\d+\.\d+", r.stdout)
                if m:
                    return int(m.group(1))
    except Exception:
        pass
    return None


# --------------------------------------------------------------------------- #
# Browser discovery (Edge preferred - far fewer CAPTCHAs than Chrome)
# --------------------------------------------------------------------------- #
_PF   = os.environ.get("PROGRAMFILES", r"C:\Program Files")
_PFX  = os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")
_LAD  = os.environ.get("LOCALAPPDATA", "")

BROWSER_PATHS = {
    "edge": [
        os.path.join(_PF,  "Microsoft", "Edge", "Application", "msedge.exe"),
        os.path.join(_PFX, "Microsoft", "Edge", "Application", "msedge.exe"),
        os.path.join(_LAD, "Microsoft", "Edge", "Application", "msedge.exe"),
    ],
    "chrome": [
        os.path.join(_PF,  "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(_PFX, "Google", "Chrome", "Application", "chrome.exe"),
        os.path.join(_LAD, "Google", "Chrome", "Application", "chrome.exe"),
    ],
}


def find_browser_binary(pref="auto"):
    """Return (path, type). pref: 'auto' | 'edge' | 'chrome'.
    'auto' tries Edge first (fewer CAPTCHAs), then Chrome."""
    pref = (pref or "auto").lower()
    order = ["edge", "chrome"] if pref in ("auto", "") else [pref, "edge", "chrome"]
    seen = []
    for bt in order:
        if bt in seen:
            continue
        seen.append(bt)
        for p in BROWSER_PATHS.get(bt, []):
            if p and os.path.isfile(p):
                return p, bt
    return None, None


def fetch_free_proxy(country_code="us", logger=print):
    """Fetch + verify a free public proxy for a country. Returns dict or None."""
    import urllib.request
    cc = (country_code or "us").upper()
    apis = [
        f"https://api.proxyscrape.com/v2/?request=displayproxies&protocol=http&timeout=5000&country={cc}&ssl=all&anonymity=all",
        f"https://www.proxy-list.download/api/v1/get?type=http&anon=elite&country={cc}",
    ]
    for api in apis:
        try:
            req = urllib.request.Request(api, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                text = resp.read().decode("utf-8", "ignore").strip()
            lines = [l.strip() for l in text.splitlines() if l.strip() and ":" in l]
            random.shuffle(lines)
            for line in lines[:5]:
                parts = line.split(":")
                if len(parts) == 2:
                    host, port = parts
                    try:
                        ph = urllib.request.ProxyHandler({"https": f"http://{host}:{port}"})
                        opener = urllib.request.build_opener(ph)
                        with opener.open("https://httpbin.org/ip", timeout=8) as tr:
                            if tr.status == 200:
                                logger(f"Free proxy found: {host}:{port} ({cc})")
                                return {"host": host, "port": port, "type": "http"}
                    except Exception:
                        continue
        except Exception:
            continue
    logger(f"No working free proxy for {cc} - continuing without one")
    return None


# --------------------------------------------------------------------------- #
# Driver factory (Edge via native Selenium, Chrome via undetected-chromedriver)
# --------------------------------------------------------------------------- #
def _common_args(profile_dir, headless, proxy, extra_extensions, lang="en"):
    """Arguments shared by Edge and Chrome. Returns list of CLI args."""
    vw, vh = random.choice(_VIEWPORTS)
    args = [
        f"--user-data-dir={profile_dir}",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        f"--lang={lang}",
        f"--window-size={vw},{vh}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-notifications",
        "--disable-popup-blocking",
        "--disable-blink-features=AutomationControlled",
        "--mute-audio",
    ]

    # Extensions (Buster, VPN) + proxy-auth helper need a visible browser.
    ext_dirs = []
    if not headless:
        auth_ext = _build_proxy_auth_extension(proxy)
        if auth_ext:
            outdir = auth_ext[:-4] + "_unpacked"
            try:
                with zipfile.ZipFile(auth_ext) as zf:
                    zf.extractall(outdir)
                ext_dirs.append(outdir)
            except Exception:
                pass
        for e in (extra_extensions or []):
            if e and os.path.isdir(e):
                ext_dirs.append(e)
    if ext_dirs:
        args.append("--load-extension=" + ",".join(ext_dirs))

    if headless:
        # extensions can't run headless; keep the window off-screen instead so
        # the (rare) Buster solve still works.
        args.append("--window-position=-32000,-32000")

    if proxy and proxy.get("host") and proxy.get("port"):
        ptype = proxy.get("type", "http")
        args.append(f"--proxy-server={ptype}://{proxy['host']}:{proxy['port']}")
    return args


def _apply_stealth(driver, country, latitude=None, longitude=None):
    try:
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument",
                               {"source": _STEALTH_JS})
    except Exception:
        pass
    try:
        driver.execute_cdp_cmd("Emulation.setTimezoneOverride",
                               {"timezoneId": country_timezone(country)})
    except Exception:
        pass
    try:
        driver.execute_cdp_cmd("Network.setExtraHTTPHeaders",
                               {"headers": {"Accept-Language": "en-US,en;q=0.9"}})
    except Exception:
        pass
    if latitude is not None and longitude is not None:
        # Grant geolocation permission to Google so it uses our overridden location
        dom = google_domain(country)
        for origin in [f"https://www.{dom}", "https://www.google.com"]:
            try:
                driver.execute_cdp_cmd("Browser.grantPermissions", {
                    "permissions": ["geolocation"],
                    "origin": origin
                })
            except Exception:
                pass
    set_geolocation(driver, country, latitude, longitude)


def _build_edge_driver(args, country, binary, logger, latitude=None, longitude=None):
    """Edge via Selenium's native WebDriver (msedgedriver auto-resolved)."""
    from selenium.webdriver import Edge, EdgeOptions
    opts = EdgeOptions()
    opts.use_chromium = True
    for a in args:
        opts.add_argument(a)
    ev = random.randint(120, 132)
    opts.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        f"AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{ev}.0.0.0 "
        f"Safari/537.36 Edg/{ev}.0.0.0")
    # Native Edge (unlike uc) accepts these - they hide the automation banner.
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    # Block all downloads and mute audio
    prefs = {
        "download_restrictions": 3,
        "download.default_directory": "NUL",
        "download.prompt_for_download": False,
        "profile.default_content_setting_values.automatic_downloads": 2,
        "profile.default_content_setting_values.notifications": 2,
    }
    opts.add_experimental_option("prefs", prefs)
    opts.add_argument("--autoplay-policy=no-user-gesture-required")
    opts.add_argument("--mute-audio")
    if binary:
        opts.binary_location = binary
    logger("Launching Edge browser...")
    driver = Edge(options=opts)
    _apply_stealth(driver, country, latitude, longitude)
    _block_downloads(driver)
    driver.set_page_load_timeout(45)
    return driver


def _build_chrome_driver(args, country, binary, logger, latitude=None, longitude=None):
    """Chrome via undetected-chromedriver."""
    import undetected_chromedriver as uc
    options = uc.ChromeOptions()
    for a in args:
        options.add_argument(a)
    if binary:
        options.binary_location = binary
    chrome_ver = get_chrome_major_version()
    kwargs = {"options": options, "use_subprocess": True}
    if chrome_ver:
        logger(f"Detected Chrome version: {chrome_ver}")
        kwargs["version_main"] = chrome_ver
    driver = uc.Chrome(**kwargs)
    _apply_stealth(driver, country, latitude, longitude)
    _block_downloads(driver)
    driver.set_page_load_timeout(45)
    return driver


def build_driver(profile_dir, proxy=None, headless=False, country="us",
                 extra_extensions=None, logger=print, browser_pref="auto",
                 latitude=None, longitude=None, lang="en"):
    """Create a hardened driver. Edge preferred (fewer CAPTCHAs); falls back to
    Chrome. browser_pref: 'auto' | 'edge' | 'chrome'."""
    binary, btype = find_browser_binary(browser_pref)
    label = {"edge": "Edge", "chrome": "Chrome"}.get(btype, "browser")
    if binary:
        logger(f"Using {label}: {binary}")
    else:
        btype = "chrome"

    args = _common_args(profile_dir, headless, proxy, extra_extensions, lang)

    if btype == "edge":
        try:
            return _build_edge_driver(args, country, binary, logger, latitude, longitude)
        except Exception as e:
            logger(f"Edge launch failed ({e}); falling back to Chrome...")
            cbin, _ = find_browser_binary("chrome")
            return _build_chrome_driver(args, country, cbin, logger, latitude, longitude)
    return _build_chrome_driver(args, country, binary, logger, latitude, longitude)


def _block_downloads(driver):
    """Disable all file downloads via CDP."""
    try:
        driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
            "behavior": "deny"
        })
    except Exception:
        pass


def is_alive(driver):
    if driver is None:
        return False
    try:
        _ = driver.current_url
        return True
    except Exception:
        return False


# --------------------------------------------------------------------------- #
# Human-like helpers
# --------------------------------------------------------------------------- #
def human_pause(a=0.6, b=1.6):
    time.sleep(random.uniform(a, b))


def human_scroll(driver, steps=None):
    """Scroll down in irregular human steps."""
    steps = steps or random.randint(3, 6)
    for _ in range(steps):
        if not is_alive(driver):
            return
        dy = random.randint(280, 620)
        try:
            driver.execute_script(f"window.scrollBy(0, {dy});")
        except Exception:
            return
        time.sleep(random.uniform(0.4, 1.1))


def human_mouse(driver):
    from selenium.webdriver.common.action_chains import ActionChains
    try:
        body = driver.find_element("tag name", "body")
        ac = ActionChains(driver)
        for _ in range(random.randint(1, 3)):
            ac.move_to_element_with_offset(
                body, random.randint(5, 400), random.randint(5, 300))
            ac.pause(random.uniform(0.1, 0.4))
        ac.perform()
    except Exception:
        pass


def human_type(element, text):
    for ch in text:
        element.send_keys(ch)
        time.sleep(random.uniform(0.04, 0.18))


def safe_get(driver, url):
    if not is_alive(driver):
        raise BrowserClosedError("Browser was closed")
    try:
        driver.get(url)
    except Exception:
        if not is_alive(driver):
            raise BrowserClosedError("Browser was closed")
        raise


def page_source(driver):
    if not is_alive(driver):
        raise BrowserClosedError("Browser was closed")
    try:
        return driver.page_source
    except Exception:
        if not is_alive(driver):
            raise BrowserClosedError("Browser was closed")
        raise


# --------------------------------------------------------------------------- #
# Consent / region warm-up
# --------------------------------------------------------------------------- #
def accept_consent(driver, logger=print):
    from selenium.webdriver.common.by import By
    selectors = [
        "button#L2AGLb", "button[aria-label*='Accept all']",
        "button[aria-label*='Accept']", "form[action*='consent'] button",
        "button.tHlp8d", "div[role='none'] button + button",
    ]
    for sel in selectors:
        try:
            for btn in driver.find_elements(By.CSS_SELECTOR, sel):
                try:
                    btn.click()
                    logger("Accepted Google consent")
                    human_pause(0.6, 1.2)
                    return True
                except Exception:
                    continue
        except Exception:
            continue
    return False


def warm_up(driver, country, logger=print, lang="en"):
    """Visit the country's Google homepage, settle cookies + consent.
    This makes the later typed search look like a continuing session."""
    dom = google_domain(country)
    logger(f"Warming up session on {dom} (region {country})...")
    safe_get(driver, f"https://www.{dom}/?gl={country}&hl={lang}")
    human_pause(1.5, 3.0)
    if not is_alive(driver):
        raise BrowserClosedError("Browser closed during warm-up")
    accept_consent(driver, logger)
    accept_google_consent(driver, logger)
    human_mouse(driver)
    human_pause(0.8, 1.8)


# --------------------------------------------------------------------------- #
# Block detection
# --------------------------------------------------------------------------- #
def classify_page(src: str) -> str:
    """Return one of: ok | consent | captcha | soft_block | http_403 | empty."""
    s = (src or "").lower()
    if not s:
        return "empty"
    # Consent / cookie wall detection (must check BEFORE "ok")
    if ("consent.google" in s or "before you continue" in s
            or 'action="https://consent.google' in s
            or 'id="cnsw"' in s or "i agree" in s):
        return "consent"
    # Real results present -> ok (even if recaptcha scripts load in bg)
    if 'id="search"' in s or 'id="rso"' in s or 'jsname="uwcknb"' in s:
        return "ok"
    if "403. that" in s or "does not have permission" in s or "error 403" in s:
        return "http_403"
    if 'id="captcha-form"' in s or "g-recaptcha" in s or "/sorry/" in s:
        return "captcha"
    if ("unusual traffic" in s or "automated queries" in s
            or "our systems have detected" in s):
        return "soft_block"
    return "empty"


def accept_google_consent(driver, logger=print):
    """Click through Google's consent/cookie page if present."""
    from selenium.webdriver.common.by import By
    try:
        src = (driver.page_source or "").lower()
        if "consent.google" not in src and "before you continue" not in src:
            return False
        logger("Google consent page detected - accepting...")
        # Try various accept button selectors
        for sel in ['button[aria-label*="Accept"]', 'button[aria-label*="accept"]',
                     '#L2AGLb', 'button[jsname="b3VHJd"]',
                     'form[action*="consent"] button', 'input[type="submit"]']:
            try:
                btns = driver.find_elements(By.CSS_SELECTOR, sel)
                for btn in btns:
                    if btn.is_displayed():
                        btn.click()
                        time.sleep(2)
                        logger("Consent accepted.")
                        return True
            except Exception:
                continue
        # Fallback: click any visible button with "agree" or "accept" text
        for btn in driver.find_elements(By.TAG_NAME, "button"):
            try:
                txt = btn.text.lower()
                if btn.is_displayed() and ("accept" in txt or "agree" in txt or "i agree" in txt):
                    btn.click()
                    time.sleep(2)
                    logger("Consent accepted (text match).")
                    return True
            except Exception:
                continue
        logger("Consent page found but could not auto-accept.")
        return False
    except Exception:
        return False


# --------------------------------------------------------------------------- #
# Human search + continuous-scroll pagination
# --------------------------------------------------------------------------- #
def human_search(driver, keyword, country, logger=print, city=None, lang="en"):
    """Type the keyword into the live Google search box and submit.
    When a city is selected, navigates via URL with UULE so Google's server
    localises results to that city. Falls back to URL nav if box not found."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from urllib.parse import quote_plus  # used in fallback URL

    if not is_alive(driver):
        raise BrowserClosedError("Browser closed before search")

    # Apply geolocation CDP override for JS-level location APIs
    if city and city in CITY_COORDS:
        lat, lng = CITY_COORDS[city]
        set_geolocation(driver, latitude=lat, longitude=lng)

    dom = google_domain(country)

    def _find_box():
        for sel in ["textarea[name='q']", "input[name='q']", "textarea#APjFqb"]:
            try:
                els = driver.find_elements(By.CSS_SELECTOR, sel)
                if els and els[0].is_displayed():
                    return els[0]
            except Exception:
                continue
        return None

    def _dismiss_consent():
        """Click through Google's cookie consent page if it appears."""
        try:
            src = driver.page_source or ""
            if "consent" not in src.lower():
                return
            for sel in [
                "button[id*='accept']", "button[aria-label*='Accept']",
                "button[aria-label*='accept']", "form[action*='consent'] button",
                "#L2AGLb", ".tHlp8d",
            ]:
                btns = driver.find_elements(By.CSS_SELECTOR, sel)
                if btns:
                    btns[0].click()
                    human_pause(1.0, 2.0)
                    return
        except Exception:
            pass

    # Make sure we're on a Google page with a search box
    cur = ""
    try:
        cur = driver.current_url or ""
    except Exception:
        pass
    if "google." not in cur:
        try:
            safe_get(driver, f"https://www.{dom}/")
            human_pause(1.5, 2.5)
        except Exception:
            warm_up(driver, country, logger)

    # Try to find search box; dismiss consent page if blocking it
    box = _find_box()
    if box is None:
        _dismiss_consent()
        human_pause(0.8, 1.5)
        box = _find_box()

    # Still not found - reload Google homepage and try once more
    if box is None:
        try:
            safe_get(driver, f"https://www.{dom}/")
            human_pause(2.0, 3.0)
            _dismiss_consent()
            human_pause(0.8, 1.5)
            box = _find_box()
        except Exception:
            pass

    if box is not None:
        try:
            box.click()
            human_pause(0.3, 0.8)
            box.send_keys(Keys.CONTROL + "a")
            human_pause(0.1, 0.2)
            box.send_keys(Keys.DELETE)
            human_pause(0.2, 0.4)
        except Exception:
            try:
                box.clear()
            except Exception:
                pass
        human_type(box, keyword)
        human_pause(0.4, 1.0)

        try:
            box.send_keys(Keys.ENTER)
        except Exception:
            box.submit()
        logger(f"Typed & submitted: '{keyword}'" + (f" (city: {city})" if city else ""))
        human_pause(2.0, 3.5)
        return True

    # Fallback: go to homepage first (sets referer + session), then search URL
    logger("Search box not found - using URL fallback via homepage")
    try:
        cur2 = driver.current_url or ""
        if "google." not in cur2:
            safe_get(driver, f"https://www.{dom}/")
            human_pause(1.5, 2.5)
    except Exception:
        pass
    safe_get(driver, f"https://www.{dom}/search?q={quote_plus(keyword)}&gl={country}&hl={lang}")
    human_pause(2.0, 3.5)
    return True


def load_more_results(driver, target_count, max_pages, logger=print):
    """Continuous-scroll: scroll + click 'More results' until we have enough
    organic results or Google stops giving more.
    Returns accumulated list of unique links across all pages."""
    from selenium.webdriver.common.by import By

    all_links = []
    seen_urls = set()

    def _accumulate():
        for link in extract_organic(driver):
            if link not in seen_urls:
                seen_urls.add(link)
                all_links.append(link)

    _accumulate()
    if len(all_links) >= target_count:
        return all_links

    stagnant = 0
    prev_count = len(all_links)
    for _ in range(max_pages + 2):
        if not is_alive(driver):
            raise BrowserClosedError("Browser closed during pagination")

        human_scroll(driver, steps=random.randint(2, 4))
        human_pause(0.5, 1.0)

        _accumulate()
        if len(all_links) >= target_count:
            return all_links
        if len(all_links) <= prev_count:
            stagnant += 1
        else:
            stagnant = 0
        prev_count = len(all_links)
        if stagnant >= 2:
            return all_links

        # Try the "More results" button (continuous scroll trigger)
        clicked = False
        # Only use continuous-scroll selectors (NOT #pnnext which navigates away)
        _accumulate()
        for sel in ["a[aria-label*='More results']",
                    "div[role='button'][aria-label*='More']",
                    "span.RVQdVd"]:
            try:
                for el in driver.find_elements(By.CSS_SELECTOR, sel):
                    try:
                        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                        human_pause(0.3, 0.7)
                        el.click()
                        clicked = True
                        logger("Clicked 'More results'")
                        human_pause(1.0, 1.8)
                        break
                    except Exception:
                        continue
            except Exception:
                continue
            if clicked:
                break
        if not clicked:
            # Last resort: traditional "Next" page link (full navigation)
            try:
                nxt = driver.find_elements(By.CSS_SELECTOR, "#pnnext, a[id='pnnext']")
                if nxt:
                    nxt[0].click()
                    logger("Clicked 'Next' page link")
                    human_pause(1.5, 2.5)
                    _accumulate()
                    if len(all_links) >= target_count:
                        return all_links
                    continue
            except Exception:
                pass
            return all_links
    return all_links


# --------------------------------------------------------------------------- #
# Organic result extraction - JavaScript-based (like Ctrl+F in a real browser)
# --------------------------------------------------------------------------- #

def _get_organic_links(src):
    """Parse page source and return ordered list of organic result URLs.
    Like Ctrl+F on the page - only main result title links, no sitelinks or ads."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(src, "html.parser")

    # Remove ads and non-organic blocks
    for sel in ["#tads", "#tadsb", ".ads-ad", "#botstuff", "#rhs",
                ".kp-wholepage", ".kp-blk", "g-scrolling-carousel"]:
        for el in soup.select(sel):
            el.decompose()

    links, seen = [], set()

    def add(href):
        if not href or not href.startswith("http"):
            return
        low = href.lower()
        if "google." in low or "/search?" in low or "webcache" in low:
            return
        if href in seen:
            return
        seen.add(href)
        links.append(href)

    rso = soup.select_one("#rso") or soup.select_one("#search") or soup

    # 1. jsname="UWckNb" - Google's primary title anchor attribute
    for a in rso.select('a[jsname="UWckNb"]'):
        add(a.get("href", ""))

    # 2. zReHs class - alternative Google title anchor class
    if not links:
        for a in rso.select("a.zReHs"):
            add(a.get("href", ""))

    # 3. h3-parent: walk up from heading to parent <a> (works across all Google layouts)
    if not links:
        for h3 in rso.select("h3"):
            a = h3.find_parent("a") or h3.find("a")
            if a:
                add(a.get("href", ""))

    # 4. ping attribute: Google sets ping on every organic title link
    if not links:
        for a in rso.select("a[ping]"):
            # Only count if it wraps or is near an h3
            if a.find("h3") or (a.parent and a.parent.find("h3")):
                add(a.get("href", ""))

    # 5. Any <a> directly containing an <h3> (broadest fallback)
    if not links:
        for a in soup.select("a"):
            if a.find("h3"):
                add(a.get("href", ""))

    return links


def extract_organic(driver, debug=False):
    """Extract ordered organic result URLs from the current SERP (like Ctrl+F)."""
    try:
        src = driver.page_source or ""
    except Exception:
        if debug:
            return [], {}
        return []

    links = _get_organic_links(src)

    if debug:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(src, "html.parser")
        info = {
            "url": "",
            "h3count": len(soup.select("h3")),
            "rso": bool(soup.select_one("#rso")),
            "jsname_count": len(soup.select('a[jsname="UWckNb"]')),
            "zReHs_count": len(soup.select("a.zReHs")),
        }
        try:
            info["url"] = driver.current_url or ""
        except Exception:
            pass
        return links, info

    return links


def _host_is_domain(link, domain_clean):
    """True only if the RESULT's HOST is the domain (or a subdomain of it). This is
    what makes the rank count by domain, not by the domain merely appearing in the
    URL path or title - e.g. uk.trustpilot.com/review/exactprint.co.uk is Trustpilot
    ranking, NOT exactprint.co.uk, so it must not be counted as the client's result."""
    from urllib.parse import urlparse
    try:
        netloc = urlparse(link if "//" in str(link) else "http://" + str(link)).netloc
    except Exception:
        netloc = ""
    host = (netloc or "").lower().split(":")[0]
    if host.startswith("www."):
        host = host[4:]
    return bool(host) and (host == domain_clean or host.endswith("." + domain_clean))


def find_domain_in_page(driver, domain_clean, page_offset=0):
    """Ctrl+F style: find the domain in the current page's organic title links.
    Returns list of {position, serp_url} with cumulative position across pages.
    Matches on the result HOST (or subdomain), not the URL path / title."""
    try:
        src = driver.page_source or ""
    except Exception:
        return []
    links = _get_organic_links(src)
    results = []
    for i, link in enumerate(links):
        if _host_is_domain(link, domain_clean):
            results.append({"position": i + 1 + page_offset, "serp_url": link})
    return results


def match_domain(links, domain_clean):
    """All 1-based positions where the domain (as the result HOST/subdomain) appears
    in the ordered links - not where it merely appears in a URL path or title."""
    matches = []
    for i, link in enumerate(links):
        if _host_is_domain(link, domain_clean):
            matches.append({"position": i + 1, "serp_url": link})
    return matches


def clean_domain(domain: str) -> str:
    return (domain.lower()
            .replace("https://", "").replace("http://", "")
            .replace("www.", "").rstrip("/"))


def bring_browser_to_front():
    """Bring the Chrome/Edge/Chromium window to the foreground on Windows."""
    try:
        import ctypes
        import ctypes.wintypes
        user32 = ctypes.windll.user32
        titles = []
        def _enum_cb(hwnd, _):
            if user32.IsWindowVisible(hwnd):
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    buf = ctypes.create_unicode_buffer(length + 1)
                    user32.GetWindowTextW(hwnd, buf, length + 1)
                    titles.append((hwnd, buf.value))
            return True
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.wintypes.HWND, ctypes.wintypes.LPARAM)
        user32.EnumWindows(WNDENUMPROC(_enum_cb), 0)
        for hwnd, title in titles:
            if any(k in title for k in ("Chrome", "Chromium", "Google", "Edge", "edge")):
                user32.ShowWindow(hwnd, 9)   # SW_RESTORE
                user32.SetForegroundWindow(hwnd)
                break
    except Exception:
        pass


def human_visit_neutral(driver, target_domain=None):
    """Visit the target domain or a neutral site between keywords to look human."""
    sites = [
        "https://www.wikipedia.org",
        "https://www.weather.com",
        "https://news.ycombinator.com",
    ]
    if target_domain:
        td = target_domain if target_domain.startswith("http") else f"https://{target_domain}"
        sites.append(td)
    url = random.choice(sites)
    try:
        driver.get(url)
        time.sleep(random.uniform(3, 7))
    except Exception:
        pass


def human_visit_neutral_bg(driver, target_domain=None, logger=print):
    """Open a neutral site in a background tab (non-blocking).
    Returns quickly - the tab loads in parallel while ranking continues."""
    from selenium.webdriver.common.keys import Keys
    sites = [
        "https://www.wikipedia.org",
        "https://www.weather.com",
        "https://news.ycombinator.com",
    ]
    if target_domain:
        td = target_domain if target_domain.startswith("http") else f"https://{target_domain}"
        sites.append(td)
    url = random.choice(sites)
    try:
        main_handle = driver.current_window_handle
        driver.execute_script(f"window.open('{url}', '_blank');")
        handles = driver.window_handles
        if len(handles) > 1:
            new_tab = [h for h in handles if h != main_handle][-1]
            driver.switch_to.window(new_tab)
            time.sleep(random.uniform(1.5, 3.0))
            driver.close()
            driver.switch_to.window(main_handle)
    except Exception:
        try:
            driver.switch_to.window(main_handle)
        except Exception:
            pass


def check_ip_location(driver=None, logger=print, use_browser=False):
    """Check public IP. use_browser=True checks through browser (detects browser VPN)."""
    import json as _json
    if use_browser and driver and is_alive(driver):
        try:
            driver.get("https://ipinfo.io/json")
            time.sleep(2)
            body = driver.find_element("tag name", "pre").text
            data = _json.loads(body)
            ip = data.get("ip", "unknown")
            city = data.get("city", "")
            region = data.get("region", "")
            country = data.get("country", "")
            loc_str = ", ".join(filter(None, [city, region, country]))
            logger(f"Current IP: {ip} - Location: {loc_str}")
            return {"ip": ip, "city": city, "region": region, "country": country, "location": loc_str}
        except Exception:
            pass
    import urllib.request
    try:
        req = urllib.request.Request("https://ipinfo.io/json",
                                    headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = _json.loads(resp.read().decode())
            ip = data.get("ip", "unknown")
            city = data.get("city", "")
            region = data.get("region", "")
            country = data.get("country", "")
            loc_str = ", ".join(filter(None, [city, region, country]))
            logger(f"Current IP: {ip} - Location: {loc_str}")
            return {"ip": ip, "city": city, "region": region, "country": country, "location": loc_str}
    except Exception:
        logger("Could not detect IP location")
        return None


def wipe_profile_if_stale(profile_dir, max_age_hours=24):
    """Remove the browser profile if it's older than max_age_hours."""
    try:
        if os.path.isdir(profile_dir):
            mtime = os.path.getmtime(profile_dir)
            age_h = (time.time() - mtime) / 3600
            if age_h > max_age_hours:
                import shutil
                shutil.rmtree(profile_dir, ignore_errors=True)
    except Exception:
        pass
