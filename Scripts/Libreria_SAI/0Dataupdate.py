import subprocess
print("\n*******************************\n Administración SAI, CAMUNDA y Zoho\n*******************************\n")
def display_menu():
    print("\nSelect an option:")
    print("1. Download SAI data")
    print("2. Make invoice dashboards")
    print("3. Process SAI data")
    print("4. Merge INSABI data")
    print("5. Update google sheet dashboard")
    print("6. Merge Zoho")
    print("7. Exit")

def run_script(script_name):
    try:
        subprocess.run(["python", script_name], text=True, check=True)
    except subprocess.CalledProcessError as e:
        print(f"Script {script_name} exited with error status {e.returncode}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

def main():
    while True:
        display_menu()
        
        choice = input("Enter your choice (1-7): ")
        
        if choice == "1":
            run_script("0web.py")
        elif choice == "2":
            run_script("0tigre.py")
        elif choice == "3":
            run_script("1cleanimss.py")
        elif choice == "4":
            run_script("2appendinsabi.py")
        elif choice == "5":
            run_script("3subeGS.py")
        elif choice == "6":
            run_script("4Zoho.py")
        elif choice == "7":
            print("Exiting the program...")
            break
        else:
            print("Invalid choice. Please select a valid option.")

if __name__ == "__main__":
    main()
