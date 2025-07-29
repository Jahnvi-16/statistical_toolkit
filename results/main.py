import os

def run_descriptive_statistics():
    os.system("python descriptive_statistics.py")

def run_experiment_design():
    os.system("python experiment_design.py")

def run_quality_control_charts():
    os.system("python quality_control_charts.py")

def main():
    while True:
        print("\n📊 Statistical Toolkit Menu")
        print("1. Descriptive Statistics")
        print("2. Experimental Design (CRD / RBD)")
        print("3. Quality Control Charts (X̄-R, X̄-S, p, np, c)")
        print("4. Exit")

        choice = input("Enter your choice (1–4): ").strip()

        if choice == "1":
            run_descriptive_statistics()
        elif choice == "2":
            run_experiment_design()
        elif choice == "3":
            run_quality_control_charts()
        elif choice == "4":
            print("👋 Exiting the toolkit. Goodbye!")
            break
        else:
            print("❌ Invalid choice. Please enter 1, 2, 3 or 4.")

if __name__ == "__main__":
    main()
