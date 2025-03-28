import time
import pandas as pd

# Load the existing data from the Excel file
def load_data(file_path):
    try:
        df = pd.read_excel(file_path)
        # Ensure 'Account Number' is treated as string
        df['Account Number'] = df['Account Number'].astype(str)
        return df
    except FileNotFoundError:
        # If file does not exist, create a new DataFrame with columns
        return pd.DataFrame(columns=["Name", "Account Number", "Pin Number", "Balance"])

# Save the data to the Excel file
def save_data(data, file_path):
    data.to_excel(file_path, index=False)

# Function to display the main menu
def display_menu():
    print("""
    Welcome to the ATM:
    1. Check Balance
    2. Withdraw Funds
    3. Deposit Funds
    4. Exit
    """)

# Function to simulate card processing
def card_processing():
    print("Please insert your CARD...")
    time.sleep(2)  # Simulate card processing delay

# Function to validate the PIN
def validate_pin(df, account_number):
    # Ensure account_number is a string for comparison
    account_number = str(account_number).strip()
    for i in range(3):  # Give the user 3 attempts to enter the correct pin
        pin = int(input("Enter your ATM PIN: "))
        # Extract PIN from DataFrame
        pin_from_df = df.loc[df["Account Number"] == account_number, "Pin Number"]
        if not pin_from_df.empty and pin_from_df.values[0] == pin:
            return True
        else:
            print("Incorrect PIN. Try again.")
    return False

# Function to handle balance inquiry
def check_balance(df, account_number):
    account_number = str(account_number).strip()
    balance = df.loc[df["Account Number"] == account_number, "Balance"]
    if not balance.empty:
        print(f"Your current balance is: ${balance.values[0]}")
    else:
        print("Account not found.")

# Function to handle withdrawal
def withdraw_funds(df, account_number):
    try:
        account_number = str(account_number).strip()
        withdraw_amount = float(input("Enter the amount to withdraw: $"))
        balance = df.loc[df["Account Number"] == account_number, "Balance"]
        if not balance.empty:
            current_balance = balance.values[0]
            if withdraw_amount <= 0:
                print("Please enter a valid amount.")
            elif withdraw_amount > current_balance:
                print("Insufficient funds.")
            else:
                new_balance = current_balance - withdraw_amount
                df.loc[df["Account Number"] == account_number, "Balance"] = new_balance
                print(f"${withdraw_amount} has been withdrawn. Your new balance is: ${new_balance}")
        else:
            print("Account not found.")
    except ValueError:
        print("Invalid input. Please enter a numeric value.")
    return df

# Function to handle deposits
def deposit_funds(df, account_number):
    try:
        account_number = str(account_number).strip()
        deposit_amount = float(input("Enter the amount to deposit: $"))
        if deposit_amount <= 0:
            print("Please enter a valid amount.")
        else:
            balance = df.loc[df["Account Number"] == account_number, "Balance"]
            if not balance.empty:
                current_balance = balance.values[0]
                new_balance = current_balance + deposit_amount
                df.loc[df["Account Number"] == account_number, "Balance"] = new_balance
                print(f"${deposit_amount} has been deposited. Your new balance is: ${new_balance}")
            else:
                print("Account not found.")
    except ValueError:
        print("Invalid input. Please enter a numeric value.")
    return df

# Function to create a new account
def create_account(df):
    name = input("Enter your name: ")
    account_number = input("Enter a new account number: ").strip()
    pin_number = int(input("Enter a new PIN number: "))
    balance = float(input("Enter an initial balance: $"))
    
    new_account = pd.DataFrame({
        "Name": [name],
        "Account Number": [account_number],
        "Pin Number": [pin_number],
        "Balance": [balance]
    })
    
    df = pd.concat([df, new_account], ignore_index=True)
    save_data(df, 'client.xlsx')
    return df

# Main function to run the ATM system
def atm_machine():
    file_path = 'client.xlsx'
    df = load_data(file_path)
    
    card_processing()
    
    existing_customer = input("Are you an existing customer? (yes/no): ").strip().lower()
    
    if existing_customer == "yes":
        account_number = input("Enter your account number: ").strip()
        if account_number not in df["Account Number"].values:
            print("Account not found.")
            return
        
        if not validate_pin(df, account_number):
            print("Too many incorrect attempts. Exiting.")
            return
        
    else:
        df = create_account(df)
        account_number = df.loc[df["Name"] == input("Enter your name: "), "Account Number"].values[0]
        pin_number = df.loc[df["Account Number"] == account_number, "Pin Number"].values[0]
        print(f"Account created successfully. Your account number is {account_number} and your PIN is {pin_number}.")
    
    while True:
        display_menu()
        
        try:
            option = int(input("Please select an option: "))
        except ValueError:
            print("Invalid input. Please enter a number.")
            continue
        
        if option == 1:
            check_balance(df, account_number)
        elif option == 2:
            df = withdraw_funds(df, account_number)
        elif option == 3:
            df = deposit_funds(df, account_number)
        elif option == 4:
            print("Thank you for using the ATM. Goodbye!")
            break
        else:
            print("Invalid option. Please select from the menu.")
    
    save_data(df, file_path)

# Start the ATM system
if _name_ == "_main_":
    atm_machine()
