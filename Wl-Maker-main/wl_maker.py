import pandas as pd
import glob
import os
import time
from color import color


def save_to_csv(raw_name, df):
    file_name = raw_name.replace("/", "_")
    os.makedirs("csvki", exist_ok=True)
    save_path = os.path.join("csvki", f"{file_name}.csv")
    df.to_csv(save_path, sep=";", index=False)
    print("\nZapisano plik! ♿")
    return file_name


gr = color.GREEN
bd = color.BOLD
eol = color.END
cn = color.DARKCYAN

print(
    f"{gr}{bd}\nWitaj kolego, słuchaj się grzecznie poleceń bo inaczej zepsujesz a po co!\n{eol}"
)
time.sleep(1)

while True:
    path = glob.glob(os.path.join(os.getcwd(), "zestawienia", "*.xlsx"))
    for i, item in enumerate(path):
        print(f'index: {i} plik: "{os.path.basename(item)}"')

    ind = int(input("Wybierz indeks ścieżki:\n"))
    path = path[ind]
    df = pd.read_excel(path)
    file_name = str(
        f"{df['Tok - nazwa'][0]} {df['Grupa - nazwa'][0]} {(df['Prowadzący zajęcia, imię'][0])[0]}{df['Prowadzący zajęcia, nazwisko'][0]}"
    )
    menu_1 = input(
        f"\nCo chcesz zrobić słodki książe?\n{gr}{bd}'q'{eol} - wygenerować CSV z gotowego pliku XLSX\n"
        f"{gr}{bd}'r'{eol} - uzupełnić zestawienie wyciągnięte z dziekanatu\n{gr}{bd}enter{eol} - zamknij program\n"
    )
    if menu_1.lower() == "q":
        df_copy = df.copy()
        df_copy = df_copy.loc[
            :,
            [
                "Imię",
                "Nazwisko",
                "Company",
                "e-mail",
                "Start Date",
                "End Date",
                "Timezone ID",
                "Trainer",
            ],
        ]
        print("zapisuje do csv\n")

        save_to_csv(file_name, df_copy)

    elif menu_1.lower() == "r":
        while True:
            df_copy = df.copy()
            teacher_mail = input(
                f"Wybierz typ maila wykładowcy\n{gr}{bd}'q'{eol} - {(df_copy['Prowadzący zajęcia, imię'][0]).lower()[0]}{df_copy['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl\n"
                f"{gr}{bd}'r'{eol} - {(df_copy['Prowadzący zajęcia, imię'][0].lower())}.{df_copy['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl\n"
                f"{gr}{bd}Jeśli mail jest niestandardowany, uzupełnij pole: {eol}\n"
            )
            if teacher_mail.lower() == "q":
                teacher_mail = f"{(df_copy['Prowadzący zajęcia, imię'][0]).lower()[0]}{df_copy['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl"
            elif teacher_mail.lower() == "r":
                teacher_mail = f"{(df_copy['Prowadzący zajęcia, imię'][0].lower())}.{df_copy['Prowadzący zajęcia, nazwisko'][0].lower()}@wsb.edu.pl"

            df_copy.loc[len(df_copy)] = {
                "Imię": df_copy["Prowadzący zajęcia, imię"][0],
                "Nazwisko": df_copy["Prowadzący zajęcia, nazwisko"][0],
                "e-mail": teacher_mail,
            }

            df_copy["Company"] = "AWSB"
            df_copy["Start Date"] = pd.Timestamp.now().normalize()
            df_copy["End Date"] = df_copy["Start Date"] + pd.Timedelta(days=14)
            df_copy["Timezone ID"] = 54
            df_copy["Trainer"] = False
            df_copy = df_copy.loc[
                :,
                [
                    "Imię",
                    "Nazwisko",
                    "Company",
                    "e-mail",
                    "Start Date",
                    "End Date",
                    "Timezone ID",
                    "Trainer",
                ],
            ]
            df_copy.loc[df_copy.index[-1], "Trainer"] = True

            if (
                input(
                    f"Sprawdź podsumowanie bratku czy jest okej.\n{gr}{bd}'enter' - okej{eol}\n{gr}{bd}'r' - nie okej{eol}\n"
                    f"\n{cn}{bd}ostatni index:{eol} {ind},{cn}{bd} przedmiot:{eol} {df['Przedmiot'][0]}\n\n{df_copy.tail()}\n"
                )
                == "r"
            ):
                continue
            else:
                save_to_csv(file_name, df_copy)
                break
    else:
        break
    action = input(
        f"\nNo i wariacie co robimy?\n{gr}{bd}'q'{eol} - Jeśli chcesz ponownie skorzystać\n{gr}{bd}'r'{eol} -  Jeśli chcesz zakończyć program\n"
    )
    if action.lower() == "r":
        break
    elif action.lower() == "q":
        continue
