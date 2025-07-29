import pandas as pd
import os
import scipy.stats as stats
import statsmodels.api as sm
from statsmodels.formula.api import ols

def convert_table_to_long(df):
    """
    Converts a wide-format DataFrame with blocks as rows and treatments as columns
    to long format with 'Block', 'Treatment', and 'Response'.
    """
    df_long = df.set_index(df.columns[0]).reset_index()
    df_long = df_long.melt(id_vars=df_long.columns[0], var_name='Treatment', value_name='Response')
    df_long.columns = ['Block', 'Treatment', 'Response']
    return df_long

def analyze_crd_table_format(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        df_long = convert_table_to_long(df)

        print("\n📊 ANOVA for CRD (Table Format)")
        model = ols('Response ~ C(Treatment)', data=df_long).fit()
        anova_table = sm.stats.anova_lm(model, typ=2)

        print(anova_table)

        p_value = anova_table['PR(>F)']['C(Treatment)']
        print(f"\n🔍 P-value = {p_value:.4f}")
        if p_value < 0.05:
            print("🚨 Reject null hypothesis: Treatments have significant differences.")
        else:
            print("✅ Fail to reject null hypothesis: No significant difference between treatments.")

    except Exception as e:
        print(f"❌ Error in CRD analysis: {e}")


def analyze_rbd_table_format(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        if df.shape[1] < 2:
            print("❌ The table must have at least one block and two treatments.")
            return

        print(f"\n📊 ANOVA for Block vs Treatment (RBD Table Format)")

        df_long = convert_table_to_long(df)
        print("\n📄 Converted Data (Long Format):")
        print(df_long)

        # Run ANOVA
        model = ols('Response ~ C(Treatment) + C(Block)', data=df_long).fit()
        anova_table = sm.stats.anova_lm(model, typ=2)

        print("\n📊 ANOVA Table:")
        print(anova_table)

        p_value = anova_table['PR(>F)']['C(Treatment)']
        print(f"\n🔍 P-value (Treatment) = {p_value:.4f}")
        if p_value < 0.05:
            print("🚨 Reject null hypothesis: Significant difference exists between treatments.")
        else:
            print("✅ Fail to reject null hypothesis: No significant difference between treatments.")

    except Exception as e:
        print(f"❌ Error: {e}")

def analyze_rbd_table_format(file_path, sheet_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        df_long = convert_table_to_long(df)

        print("\n📊 ANOVA for RBD (Table Format)")
        model = ols('Response ~ C(Treatment) + C(Block)', data=df_long).fit()
        anova_table = sm.stats.anova_lm(model, typ=2)

        print(anova_table)

        p_value = anova_table['PR(>F)']['C(Treatment)']
        print(f"\n🔍 P-value (Treatment) = {p_value:.4f}")
        if p_value < 0.05:
            print("🚨 Reject null hypothesis: Treatments have significant differences.")
        else:
            print("✅ Fail to reject null hypothesis: No significant difference between treatments.")

    except Exception as e:
        print(f"❌ Error in RBD analysis: {e}")

if __name__ == "__main__":
    print("📊 Experimental Design ANOVA (Excel Table Format)")
    print("Choose design type:")
    print("1. Completely Randomized Design (CRD)")
    print("2. Randomized Block Design (RBD)")

    choice = input("Enter 1 or 2: ").strip()
    file_path = input("Enter Excel file path (e.g., datasets/design_data.xlsx): ").strip()
    sheet_name = input("Enter sheet name: ").strip()

    if not os.path.exists(file_path):
        print("❌ File not found.")
    elif choice == "1":
        analyze_crd_table_format(file_path, sheet_name)
    elif choice == "2":
        analyze_rbd_table_format(file_path, sheet_name)
    else:
        print("❌ Invalid choice. Please enter 1 or 2.")
