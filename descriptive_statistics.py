import statistics
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import os

def des_stats(data):
    """Calculate descriptive statistics for a list of numerical data."""
    count = len(data)
    total = sum(data)
    maxx = max(data)
    minn = min(data)
    mean = statistics.mean(data)
    Range = maxx - minn
    median = statistics.median(data)
    try:
        mode = statistics.mode(data)
    except statistics.StatisticsError:
        mode = "No unique mode"
    std_dev = statistics.stdev(data)
    variance = statistics.variance(data)

    return {
        "Count": count,
        "Total": total,
        "Max": maxx,
        "Min": minn,
        "Mean": mean,
        "Range": Range,
        "Median": median,
        "Mode": mode,
        "Std Dev": std_dev,
        "Variance": variance
    }

def read_excel_column(file_path, sheet_name, col_letter):
    """Read a column of numeric data from an Excel file using openpyxl."""
    try:
        wb = load_workbook(file_path, data_only=True)
        sheet = wb[sheet_name]
        data = []

        for cell in sheet[col_letter]:
            if isinstance(cell.value, (int, float)):
                data.append(cell.value)

        return data
    except Exception as e:
        print(f"❌ Error reading Excel: {e}")
        return []

if __name__ == "__main__":
    print("📊 Descriptive Statistics Calculator")
    print("Choose data input method:")
    print("1. Enter a list of numbers")
    print("2. Read from an Excel file")

    choice = input("Enter 1 or 2: ")

    if choice == "1":
        raw = input("Enter numbers separated by commas: ")
        try:
            data = [float(x.strip()) for x in raw.split(",")]
            stats = des_stats(data)

            print("\n✅ Calculated Descriptive Statistics:")
            for key, value in stats.items():
                print(f"{key:10}: {value}")
        except ValueError:
            print("❌ Invalid input. Please enter numeric values only.")

    elif choice == "2":
        file_path = input("Enter Excel file path (e.g., datasets/data.xlsx): ").strip()
        sheet_name = input("Enter sheet name (e.g., Sheet1): ").strip()
        col_letter = input("Enter column letter (e.g., A): ").strip().upper()

        if os.path.exists(file_path):
            data = read_excel_column(file_path, sheet_name, col_letter)
            if data:
                stats = des_stats(data)
                print("\n✅ Calculated Descriptive Statistics from Excel:")
                for key, value in stats.items():
                    print(f"{key:10}: {value}")
            else:
                print("⚠️ No numeric data found or incorrect column.")
        else:
            print("❌ File not found.")

    else:
        print("❌ Invalid choice. Please enter 1 or 2.")
