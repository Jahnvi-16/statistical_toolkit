import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np

def calculate_xbar_r(data):
    xbar = data.mean()
    R = data.max() - data.min()
    return xbar, R

def create_xbar_r_chart(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return

    if df.empty:
        print("❌ Excel sheet is empty.")
        return

    # Calculate mean and range for each column (sample)
    xbars = []
    Rs = []

    for col in df.columns:
        xbar, R = calculate_xbar_r(df[col])
        xbars.append(xbar)
        Rs.append(R)

    # Overall averages
    X_bar_bar = np.mean(xbars)
    R_bar = np.mean(Rs)

    n = df.shape[0]  # Sample size (rows)

    # Constants for X̄ and R chart (from standard control chart constants table)
    A2_values = {
        2: 1.880, 3: 1.023, 4: 0.729, 5: 0.577, 6: 0.483,
        7: 0.419, 8: 0.373, 9: 0.337, 10: 0.308, 11: 0.285,
        12: 0.266, 13: 0.249, 14: 0.235, 15: 0.223, 16: 0.212,
        17: 0.203, 18: 0.194, 19: 0.187, 20: 0.180, 21: 0.173,
        22: 0.167, 23: 0.162, 24: 0.157, 25: 0.153
    }
    D3_values = {
        2: 0, 3: 0, 4: 0, 5: 0, 6: 0,
        7: 0.076, 8: 0.136, 9: 0.184, 10: 0.223, 11: 0.256,
        12: 0.283, 13: 0.307, 14: 0.328, 15: 0.347, 16: 0.363,
        17: 0.378, 18: 0.391, 19: 0.403, 20: 0.415, 21: 0.425,
        22: 0.434, 23: 0.443, 24: 0.451, 25: 0.459
    }
    D4_values = {
        2: 3.267, 3: 2.574, 4: 2.282, 5: 2.115, 6: 2.004,
        7: 1.924, 8: 1.864, 9: 1.816, 10: 1.777, 11: 1.744,
        12: 1.717, 13: 1.693, 14: 1.672, 15: 1.653, 16: 1.637,
        17: 1.622, 18: 1.608, 19: 1.597, 20: 1.585, 21: 1.575,
        22: 1.566, 23: 1.557, 24: 1.548, 25: 1.541
    }

    if n not in A2_values:
        print("❌ Sample size not supported (currently supports n = 2 to 25).")
        return

    A2 = A2_values[n]
    D3 = D3_values[n]
    D4 = D4_values[n]

    # Control limits
    xbar_UCL = X_bar_bar + A2 * R_bar
    xbar_LCL = X_bar_bar - A2 * R_bar
    R_UCL = D4 * R_bar
    R_LCL = D3 * R_bar

    print("\n📊 X̄ Chart and R Chart Control Limits:")
    print(f"X̄-bar = {X_bar_bar:.2f}, R-bar = {R_bar:.2f}")
    print(f"X̄ Chart: UCL = {xbar_UCL:.2f}, CL = {X_bar_bar:.2f}, LCL = {xbar_LCL:.2f}")
    print(f"R Chart:  UCL = {R_UCL:.2f}, CL = {R_bar:.2f}, LCL = {R_LCL:.2f}")

    # Plot X̄ chart
    plt.figure()
    plt.plot(xbars, marker='o', label='X̄ values')
    plt.axhline(X_bar_bar, color='green', linestyle='--', label='CL')
    plt.axhline(xbar_UCL, color='red', linestyle='--', label='UCL')
    plt.axhline(xbar_LCL, color='red', linestyle='--', label='LCL')
    plt.title('X̄ Chart')
    plt.xlabel('Subgroup')
    plt.ylabel('Mean')
    plt.legend()
    plt.grid(True)
    plt.show()

    # Plot R chart
    plt.figure()
    plt.plot(Rs, marker='o', label='R values')
    plt.axhline(R_bar, color='green', linestyle='--', label='CL')
    plt.axhline(R_UCL, color='red', linestyle='--', label='UCL')
    plt.axhline(R_LCL, color='red', linestyle='--', label='LCL')
    plt.title('R Chart')
    plt.xlabel('Subgroup')
    plt.ylabel('Range')
    plt.legend()
    plt.grid(True)
    plt.show()

if __name__ == "__main__":
    print("📈 Quality Control Chart Generator - X̄ & R Chart")
    file_path = input("Enter Excel file path (e.g., datasets/variable_data.xlsx): ").strip()
    sheet_name = input("Enter sheet name (e.g., XbarR): ").strip()

    if os.path.exists(file_path):
        create_xbar_r_chart(file_path, sheet_name)
    else:
        print("❌ File not found.")
