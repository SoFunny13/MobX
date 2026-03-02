import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Reference file paths
SOURCES_FILE = os.path.join(BASE_DIR, "id_Source.xlsx")
OFFERS_FILE = os.path.join(BASE_DIR, "офферы ФМС.xlsx")

# Sheet names and data layout
SOURCES_SHEET = "Лист1"
SOURCES_DATA_START_ROW = 2
SOURCES_COL_ID = 0       # Column A
SOURCES_COL_NAME = 1     # Column B

OFFERS_SHEET = "Лист3"
OFFERS_DATA_START_ROW = 4
OFFERS_COL_ID = 0        # Column A
OFFERS_COL_NAME = 1      # Column B

# Stats file column headers (found dynamically)
STATS_OFFER_HEADER = "Оффер"
STATS_SOURCE_HEADER = "Источник"
STATS_COST_HEADER = "Расход"

# Sources where cost stays in RUB (no USD conversion)
NO_CONVERT_SOURCES = {"yandexdirect_int"}

# Upload constraints
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16 MB
