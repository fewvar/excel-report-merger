"""Создаёт папку с примерами месячных отчётов — чтобы утилиту было на чём запустить."""

import random
import sys
from pathlib import Path

from openpyxl import Workbook

MONTHS = ["январь", "февраль", "март", "апрель", "май"]
MANAGERS = ["Соколов", "Ильина", "Ковалёв", "Терентьева", "Медведев"]
PRODUCTS = [
    ("Наушники TWS", 3490),
    ("Клавиатура механическая", 6200),
    ("Монитор 27\"", 21900),
    ("Мышь беспроводная", 1850),
    ("Док-станция USB-C", 7400),
    ("SSD 1 ТБ", 8300),
]
HEADER = ["Дата", "Менеджер", "Товар", "Количество", "Цена", "Сумма"]


def make_month(path: Path, month_index: int, seed: int) -> int:
    random.seed(seed)
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Продажи"
    worksheet.append(HEADER)

    rows = random.randint(12, 20)
    for _ in range(rows):
        day = random.randint(1, 28)
        product, price = random.choice(PRODUCTS)
        quantity = random.randint(1, 6)
        worksheet.append([
            f"{day:02d}.{month_index:02d}.2026",
            random.choice(MANAGERS),
            product,
            quantity,
            price,
            quantity * price,
        ])

    workbook.save(path)
    return rows


def main() -> int:
    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("демо-отчёты")
    folder.mkdir(parents=True, exist_ok=True)

    total = 0
    for index, month in enumerate(MONTHS, start=1):
        path = folder / f"продажи-{index:02d}-{month}.xlsx"
        total += make_month(path, index, seed=index * 17)
        print(f"  + {path.name}")

    print(f"\nГотово: {len(MONTHS)} файлов, {total} строк в папке {folder}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
