from datetime import date
import pandas as pd
import logging
import argparse
import sys

# Was written by Ellie, Chodjayev - for problems please contact.

# Create logger for later use
logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s [%(levelname)s] %(message)s",
                    handlers=[
                        logging.FileHandler(filename="info.log"),
                        logging.StreamHandler(sys.stdout)
                    ])
logger = logging.getLogger()


def inventory():
    """This will check whether the memory is available in the cabinet"""
    today = date.today()
    reversed_date = today.strftime("%m-%d-%Y")
    file = "../MMEX Inventory.xlsx"
    df = pd.ExcelFile(file).sheet_names

    # Filter sheets
    counter = 0
    sheets = []
    for sheet in df:
        if sheet == "EXTRA" or sheet == "Inventory Rules" or sheet == "Removed lines" or sheet == "EOL_Hynix_SODIMM" \
                or sheet == "EV" or sheet == "LPDDR4" or sheet == "LP4":
            pass
        else:
            counter += 1
            sheets.append(sheet)

    memory = input("Select Memory : ")

    # Loop through sheets
    counter = 0
    for i in sheets:
          if counter == len(sheets) - 1:
              break
          else:
              # Read xlsx and current sheet
              df = pd.read_excel(f"{file}", f"{sheets[counter]}")

              # Compare and keep matching columns
              a = df.columns
              b = ['IDC S/N', 'ECC', 'Cabinet Qty', 'MMEX', 'VDR']
              keep_columns = [x for x in a if x in b]

              # Maximum width on output
              pd.set_option('display.max_columns', None)
              pd.set_option('display.width', None)
              pd.set_option('display.max_colwidth', None)

              # Search IDC S/N from user input
              df = df.loc[df['IDC S/N'].str.lower().str.contains(memory.lower(), na=False), keep_columns]
              df.reset_index(drop=True, inplace=True)

              # Convert from float to int
              try:
                  df['MMEX'] = df['MMEX'].astype(int)
                  df['VDR'] = df['VDR'].astype(int)
              except KeyError:
                  pass
              finally:
                  pass
              if df.empty:
                  counter += 1
              else:
                  print(f"{sheets[counter]}\n" f"{df}\n")
                  counter += 1


while True:
    inventory()
