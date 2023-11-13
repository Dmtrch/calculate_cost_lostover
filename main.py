import pandas as pd
import tkinter as tk
from tkinter import filedialog

# чтение файла
def choose_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()

    return file_path


# добавка строки в выходной датафрейм
def add_data(name, sum_leftovers, sum_weight):
    new_row = {0: name, 1: round(sum_leftovers,2), 2: round(sum_weight,3), 3: round((sum_leftovers / sum_weight),2)}
    #out_df.append(new_row, ignore_index=True)
    out_df.loc[len(out_df.index)] = new_row



selected_file = choose_file()

df = pd.read_excel(selected_file)

out_df = pd.DataFrame(columns=range(4))



del df[df.columns[0]]

df.columns = range(len(df.columns))

name_type = df.iloc[3, 0]
sum_leftovers, sum_weight = 0, 0

for index, row in df.iloc[3:].iterrows():
    #Проверка условия для каждой строки
    if pd.notna(row[4]) and row[0] != name_type:
        # Выполнение действий, если условие выполняется
        print(name_type, round(sum_leftovers,2), round(sum_weight,3) , round((sum_leftovers / sum_weight),2) )
        add_data(name_type,sum_leftovers,sum_weight)
        name_type = row[0]  # Обновление переменной name_type
        sum_leftovers , sum_weight = 0,0

    else:
        if pd.notna(row[8]):
            sum_leftovers += row[7] * row[8]
            sum_weight += row[7]

print(name_type, round(sum_leftovers,2), round(sum_weight,3) , round((sum_leftovers / sum_weight),2) )
add_data(name_type,sum_leftovers,sum_weight)

print(out_df)
out_df.to_excel('output.xlsx', index=False)