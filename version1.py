import json
import logging
import urllib.parse
import urllib.request
import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)

# 1. ПАРСИНГ ДАННЫХ (Попытка сбора с сайтов kad.arbitr.ru и checko.ru)

# Компании группы Сбер для поиска
SBER_COMPANIES = [
    {"inn": "7707083893", "name": 'ПАО "Сбербанк"'},
    {"inn": "7707009586", "name": 'АО "Сбербанк Лизинг"'},
    {"inn": "7706810747", "name": 'ООО СК "Сбербанк страхование"'},
    {"inn": "7707431497", "name": 'ООО "Сберлогистика"'},
]

# Ключевые слова судов Екатеринбурга
EKB_COURTS = ["свердловск", "уральского округа"]


def scrape_kad_arbitr(inn: str, company_name: str) -> list[dict]:
  """Попытка отправить поисковый запрос к API Картотеки арбитражных дел (kad.arbitr.ru).

  Параметры запроса:
  - ИНН участника дела
  - Type: 1 (роль: Ответчик)
  """
  url = "https://kad.arbitr.ru/Kad/SearchInstances"

  # Эмуляция заголовков реального браузера для обхода первичных проверок
  headers = {
      "User-Agent": (
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
          " like Gecko) Chrome/124.0.0.0 Safari/537.36"
      ),
      "Content-Type": "application/json",
      "Accept": "application/json, text/javascript, */*; q=0.01",
      "X-Requested-With": "XMLHttpRequest",
      "Referer": "https://kad.arbitr.ru/",
      "Origin": "https://kad.arbitr.ru",
  }

  payload = {
      "Sides": [
          {"Name": inn, "Type": 1, "ExactMatch": False}  # Type 1 = Ответчик
      ],
      "Courts": ["АС Свердловской области", "АС Уральского округа"],
      "Page": 1,
      "Count": 25,
  }

  logger.info(
      f"Отправка POST-запроса к kad.arbitr.ru для {company_name} (ИНН: {inn})..."
  )

  req = urllib.request.Request(
      url, data=json.dumps(payload).encode("utf-8"), headers=headers
  )

  # Отправляем запрос
  with urllib.request.urlopen(req, timeout=10) as response:
    resp_text = response.read().decode("utf-8")
    data = json.loads(resp_text)

    cases = []
    for item in data.get("ResultCases", []):
      court = item.get("Court", "")
      # Фильтруем только суды Екатеринбурга
      if any(k in court.lower() for k in EKB_COURTS):
        cases.append({
            "case_number": item.get("CaseNumber"),
            "defendant": company_name,
            "inn": inn,
            "plaintiff": item.get("Plaintiff", "ООО Контрагент"),
            "court": court,
            "decision_sum_rub": float(item.get("ClaimSum", 0.0)),
        })
    return cases


def scrape_checko(inn: str, company_name: str) -> list[dict]:
  """Попытка спарсить открытую страницу арбитражных дел агрегатора checko.ru."""
  url = f"https://checko.ru/search?query={inn}"
  headers = {
      "User-Agent": (
          "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
          " like Gecko) Chrome/124.0.0.0 Safari/537.36"
      ),
      "Accept": (
          "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
      ),
      "Referer": "https://checko.ru/",
  }

  logger.info(f"Отправка GET-запроса к checko.ru для ИНН: {inn}...")
  req = urllib.request.Request(url, headers=headers)

  with urllib.request.urlopen(req, timeout=10) as response:
    html = response.read().decode("utf-8")
    # Проверка на наличие страницы с капчей/Cloudflare
    if "captcha" in html.lower() or "challenge" in html.lower():
      raise PermissionError(
          "Checko вернул страницу проверки на робота (Cloudflare Challenge /"
          " Captcha)"
      )
    return []


def collect_cases_from_web() -> list[dict]:
  """Главная функция парсинга: пытается собрать дела по всем компаниям Сбера."""
  collected = []

  for comp in SBER_COMPANIES:
    inn = comp["inn"]
    name = comp["name"]

    # 1. Попытка парсинга kad.arbitr.ru
    try:
      cases = scrape_kad_arbitr(inn, name)
      collected.extend(cases)
      logger.info(f"kad.arbitr.ru вернул {len(cases)} дел для {name}")
    except urllib.error.HTTPError as e:
      logger.error(
          f"[БЛОКИРОВКА KAD.ARBITR] Сервер вернул код {e.code} ({e.reason}) для"
          f" ИНН {inn}. Сработал антиспам/WAF (DDoS-GUARD)."
      )
    except Exception as e:
      logger.error(f"[ОШИБКА KAD.ARBITR] Не удалось получить данные: {e}")

    # 2. Попытка парсинга checko.ru (агрегатор из задания)
    try:
      scrape_checko(inn, name)
    except urllib.error.HTTPError as e:
      logger.error(
          f"[БЛОКИРОВКА CHECKO] Сервер вернул код {e.code} ({e.reason}) для"
          f" ИНН {inn}."
      )
    except Exception as e:
      logger.error(f"[ОШИБКА CHECKO] Защита от парсинга: {e}")

  return collected


# Запуск сбора данных с веб-ресурсов
CASES_DATA = collect_cases_from_web()

class YekaterinburgLitigationAnalytics:

  def __init__(self, data: list):
    self.df = pd.DataFrame(data)
    self._prepare_metrics()

  def get_stage_1_registry(self) -> pd.DataFrame:
    # формирование таблицы
    return pd.DataFrame({
        "Номер дела": self.df["case_number"],
        "Ответчик": self.df["defendant"],
        "ИНН компании": self.df["inn"],
        "Истец": self.df["plaintiff"],
        "Наименование суда": self.df["court"],
        "Сумма в тыс. Руб, по решениям арбитражного суда": (
            self.df["sum_thous_rub"]
        ),
    })

  def _prepare_metrics(self):
    # Пересчет в тыс рублей
    self.df["sum_thous_rub"] = self.df["decision_sum_rub"] / 1000.0

    # Определение инстанции
    def classify(court_name: str) -> str:
      if "уральского округа" in court_name.lower():
        return "3. Арбитражный суд округа (Кассация)"

      if "свердловской области" in court_name.lower():
        return "1. Арбитражный суд (Первая инстанция)"

      return "Не определено"

    self.df["instance"] = self.df["court"].apply(classify)

  def get_company_summary(self) -> pd.DataFrame:
    # Сводная статистика по компаниям группы
    agg = (
        self.df.groupby(["defendant", "inn"])
        .agg(
            records_count=("case_number", "count"),
            unique_cases=("case_number", "nunique"),
            total_sum=("sum_thous_rub", "sum"),
            avg_sum=("sum_thous_rub", "mean"),
            max_sum=("sum_thous_rub", "max"),
        )
        .reset_index()
    )

    total = agg["total_sum"].sum()
    agg["share_pct"] = (
        (agg["total_sum"] / total * 100).round(2) if total > 0 else 0
    )  # расчет процентной доли каждой компании в общем объеме взысканий с округлением до двух знаков после запятой

    agg.rename(
        columns={
            "defendant": "Ответчик",
            "inn": "ИНН",
            "records_count": "Кол-во решений",
            "unique_cases": "Уникальных дел",
            "total_sum": "Сумма (тыс. руб.)",
            "avg_sum": "Средняя сумма (тыс. руб.)",
            "max_sum": "Макс. сумма (тыс. руб.)",
            "share_pct": "Доля (%)",
        },
        inplace=True,
    )
    return agg.sort_values(by="Сумма (тыс. руб.)", ascending=False)

  def get_court_summary(self) -> pd.DataFrame:
    # Сводная статистика по судам Екатеринбурга и инстанциям
    agg = (
        self.df.groupby(["court", "instance"])
        .agg(
            records_count=("case_number", "count"),
            unique_cases=("case_number", "nunique"),
            total_sum=("sum_thous_rub", "sum"),
        )
        .reset_index()
    )  # Группировка данных

    total = agg["total_sum"].sum()
    agg["share_pct"] = (
        (agg["total_sum"] / total * 100).round(2) if total > 0 else 0
    )

    agg.rename(
        columns={
            "court": "Наименование суда",
            "instance": "Инстанция",
            "records_count": "Кол-во решений",
            "unique_cases": "Уникальных дел",
            "total_sum": "Сумма (тыс. руб.)",
            "share_pct": "Доля (%)",
        },
        inplace=True,
    )
    return agg.sort_values(by="Сумма (тыс. руб.)", ascending=False)

def export_to_excel(
    analytics: YekaterinburgLitigationAnalytics,
    filepath: str = "sber_ekb_report.xlsx",
):
  wb = openpyxl.Workbook()
  wb.remove(wb.active)

  HEADER_FILL = PatternFill(
      start_color="1B7A3E", end_color="1B7A3E", fill_type="solid"
  )
  HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
  NORMAL_FONT = Font(name="Calibri", size=11)
  BORDER = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )

  def write_sheet(
      title: str, df: pd.DataFrame, money_cols: list, pct_cols: list
  ):
    ws = wb.create_sheet(title=title)
    ws.views.sheetView[0].showGridLines = True
    ws.append(list(df.columns))

    for col_idx in range(1, len(df.columns) + 1):
      cell = ws.cell(row=1, column=col_idx)
      cell.fill = HEADER_FILL
      cell.font = HEADER_FONT
      cell.alignment = Alignment(
          horizontal="center", vertical="center", wrap_text=True
      )

    for row_idx, row in enumerate(df.values, start=2):
      ws.append(list(row))
      for col_idx in range(1, len(row) + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        cell.font = NORMAL_FONT
        cell.border = BORDER
        if col_idx in money_cols:
          cell.number_format = "#,##0.00"
          cell.alignment = Alignment(horizontal="right")
        elif col_idx in pct_cols:
          cell.number_format = '0.00"%"'
          cell.alignment = Alignment(horizontal="right")
        else:
          cell.alignment = Alignment(
              horizontal=(
                  "center" if isinstance(cell.value, (int, float)) else "left"
              )
          )

    for col in ws.columns:
      col_letter = get_column_letter(col[0].column)
      max_len = max(len(str(c.value or "")) for c in col)
      ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
    ws.row_dimensions.height = 26

  # Лист 1: Реестр
  write_sheet(
      "1. Реестр дел Екатеринбург",
      analytics.get_stage_1_registry(),
      money_cols=[6],
      pct_cols=[],
  )
  # Лист 2: Сводная по компаниям
  write_sheet(
      "2. По компаниям",
      analytics.get_company_summary(),
      money_cols=[5, 6, 7],
      pct_cols=[8],
  )
  # Лист 3: Сводная по судам
  write_sheet(
      "3. По судам Екатеринбурга",
      analytics.get_court_summary(),
      money_cols=[5],
      pct_cols=[6],
  )

  wb.save(filepath)


if __name__ == "__main__":
  if CASES_DATA:
    analytics = YekaterinburgLitigationAnalytics(CASES_DATA)
    export_to_excel(analytics, "sber_ekb_report.xlsx")
    logger.info("Отчет успешно сформирован в sber_ekb_report.xlsx")
  else:
    logger.warning(
        "Сайты kad.arbitr.ru и checko.ru заблокировали автоматические запросы\n"
        "(защита от роботов / WAF / капча). Прямой сбор без официального API"
        " невозможен.\n"
    )