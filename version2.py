import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 1. Данные: суды Екатеринбурга, где компании сбера — ответчики

CASES_DATA = [
    {
        "case_number": "№ А40-21566/26-48-172",
        "defendant": 'ПАО "Сбербанк"',
        "inn": "7707083893",
        "plaintiff": 'ООО "ТОРГОВАЯ КОМПАНИЯ "ОРГАНИКА"',
        "court": "Арбитражный суд Свердловской области",
        "decision_sum_rub": 1097900100.0  # 1 097 900.10 тыс. руб. (из задания)
    },
    {
        "case_number": "№ А60-14289/2026",
        "defendant": 'ПАО "Сбербанк"',
        "inn": "7707083893",
        "plaintiff": 'ООО "УралСпецМаш"',
        "court": "Арбитражный суд Уральского округа",
        "decision_sum_rub": 67320000.0   # 67 320.00 тыс. руб.
    },
    {
        "case_number": "№ А60-33410/2025",
        "defendant": 'АО "Сбербанк Лизинг"',
        "inn": "7707009586",
        "plaintiff": 'ООО "ЕкатеринбургТранс"',
        "court": "Арбитражный суд Свердловской области",
        "decision_sum_rub": 42150000.0   # 42 150.00 тыс. руб.
    },
    {
        "case_number": "№ А60-33410/2025",
        "defendant": 'АО "Сбербанк Лизинг"',
        "inn": "7707009586",
        "plaintiff": 'ООО "ЕкатеринбургТранс"',
        "court": "Арбитражный суд Уральского округа",
        "decision_sum_rub": 42150000.0   # 42 150.00 тыс. руб.
    },
    {
        "case_number": "№ А60-51204/2025",
        "defendant": 'ООО СК "Сбербанк страхование"',
        "inn": "7706810747",
        "plaintiff": 'АО "Уральский завод тяжелого машиностроения"',
        "court": "Арбитражный суд Свердловской области",
        "decision_sum_rub": 18400000.0   # 18 400.00 тыс. руб.
    },
    {
        "case_number": "№ А60-51204/2025",
        "defendant": 'ООО СК "Сбербанк страхование"',
        "inn": "7706810747",
        "plaintiff": 'АО "Уральский завод тяжелого машиностроения"',
        "court": "Арбитражный суд Уральского округа",
        "decision_sum_rub": 18400000.0   # 18 400.00 тыс. руб.
    },
    {
        "case_number": "№ А60-08912/2026",
        "defendant": 'ООО "Сберлогистика"',
        "inn": "7707431497",
        "plaintiff": 'ИП Смирнов А. В.',
        "court": "Арбитражный суд Свердловской области",
        "decision_sum_rub": 1650400.0    # 1 650.40 тыс. руб.
    }
]

#2 Анализ

class YekaterinburgLitigationAnalytics:
    def __init__(self, data: list):
        self.df = pd.DataFrame(data)
        self._prepare_metrics()

    def get_stage_1_registry(self) -> pd.DataFrame:
        #формирование таблицы
        return pd.DataFrame({
            "Номер дела": self.df["case_number"],
            "Ответчик": self.df["defendant"],
            "ИНН компании": self.df["inn"],
            "Истец": self.df["plaintiff"],
            "Наименование суда": self.df["court"],
            "Сумма в тыс. Руб, по решениям арбитражного суда": self.df["sum_thous_rub"]
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
        #Сводная статистика по компаниям группы
        agg = self.df.groupby(["defendant", "inn"]).agg(
            records_count=("case_number", "count"),
            unique_cases=("case_number", "nunique"),
            total_sum=("sum_thous_rub", "sum"),
            avg_sum=("sum_thous_rub", "mean"),
            max_sum=("sum_thous_rub", "max")
        ).reset_index()

        total = agg["total_sum"].sum()
        agg["share_pct"] = (agg["total_sum"] / total * 100).round(2) if total > 0 else 0#расчет процентной доли каждой компании в общем объеме взысканий с округлением до двух знаков после запятой

        agg.rename(columns={
            "defendant": "Ответчик",
            "inn": "ИНН",
            "records_count": "Кол-во решений",
            "unique_cases": "Уникальных дел",
            "total_sum": "Сумма (тыс. руб.)",
            "avg_sum": "Средняя сумма (тыс. руб.)",
            "max_sum": "Макс. сумма (тыс. руб.)",
            "share_pct": "Доля (%)"
        }, inplace=True)
        return agg.sort_values(by="Сумма (тыс. руб.)", ascending=False)

    def get_court_summary(self) -> pd.DataFrame:
        #Сводная статистика по судам Екатеринбурга и инстанциям
        agg = self.df.groupby(["court", "instance"]).agg(
            records_count=("case_number", "count"),
            unique_cases=("case_number", "nunique"),
            total_sum=("sum_thous_rub", "sum")
        ).reset_index()#Группировка данных

        total = agg["total_sum"].sum()
        agg["share_pct"] = (agg["total_sum"] / total * 100).round(2) if total > 0 else 0

        agg.rename(columns={
            "court": "Наименование суда",
            "instance": "Инстанция",
            "records_count": "Кол-во решений",
            "unique_cases": "Уникальных дел",
            "total_sum": "Сумма (тыс. руб.)",
            "share_pct": "Доля (%)"
        }, inplace=True)
        return agg.sort_values(by="Сумма (тыс. руб.)", ascending=False)

# 3 Экспорт в  Excel

def export_to_excel(analytics: YekaterinburgLitigationAnalytics, filepath: str = "sber_ekb_report.xlsx"):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    HEADER_FILL = PatternFill(start_color="1B7A3E", end_color="1B7A3E", fill_type="solid")
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    NORMAL_FONT = Font(name="Calibri", size=11)
    BORDER = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    def write_sheet(title: str, df: pd.DataFrame, money_cols: list, pct_cols: list):
        ws = wb.create_sheet(title=title)
        ws.views.sheetView[0].showGridLines = True
        ws.append(list(df.columns))

        for col_idx in range(1, len(df.columns) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for row_idx, row in enumerate(df.values, start=2):
            ws.append(list(row))
            for col_idx in range(1, len(row) + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = NORMAL_FONT
                cell.border = BORDER
                if col_idx in money_cols:
                    cell.number_format = '#,##0.00'
                    cell.alignment = Alignment(horizontal="right")
                elif col_idx in pct_cols:
                    cell.number_format = '0.00"%"'
                    cell.alignment = Alignment(horizontal="right")
                else:
                    cell.alignment = Alignment(horizontal="center" if isinstance(cell.value, (int, float)) else "left")

        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max(len(str(c.value or "")) for c in col)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
        ws.row_dimensions.height = 26

    # Лист 1: Реестр
    write_sheet("1. Реестр дел Екатеринбург", analytics.get_stage_1_registry(), money_cols=[6], pct_cols=[])
    # Лист 2: Сводная по компаниям
    write_sheet("2. По компаниям", analytics.get_company_summary(), money_cols=[5, 6, 7], pct_cols=[8])
    # Лист 3: Сводная по судам
    write_sheet("3. По судам Екатеринбурга", analytics.get_court_summary(), money_cols=[5], pct_cols=[6])

    wb.save(filepath)

if __name__ == "__main__":
    analytics = YekaterinburgLitigationAnalytics(CASES_DATA)
    export_to_excel(analytics, "sber_ekb_report.xlsx")