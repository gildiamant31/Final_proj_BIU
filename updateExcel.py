import pandas as pd
import os
from tkinter import *
from Bio import SeqIO

# first excel file from Neta
initial_worksheet = pd.read_excel('src/ISR_Random_Nov2021.xlsx')
# update file name everytime we got new one :
if len(os.listdir('src/input/')) > 0:
    files = ["src/input/" + x for x in os.listdir('src/input') if x.endswith(".xlsx")]
    newest = max(files, key=os.path.getctime)
else:
    newest = 'src/ISR_Random_Nov2021.xlsx'
# this block of code use tkinter to ask from the user path to excel file.
window = Tk()
window.title("Get new Excel file")
main_lst = []
label1 = Label(window, text="New file name: ", padx=20, pady=10)
d = Entry(window, width=30, borderwidth=5)
d.insert(END, newest)
Exit = Button(window, text="change new file", padx=20, pady=10, command=window.quit)
label1.grid(row=0, column=0)
d.grid(row=0, column=1)
Exit.grid(row=5, column=0, columnspan=2)
window.mainloop()
newest = d.get()
window.destroy()
window.quit()
# using of tkinter gui ends here


current_worksheet = pd.read_excel(newest)
#  remove the index column - ## need better way for doing that!! ##
current_worksheet.drop(current_worksheet.columns[0], 1, inplace=True)

# check if there are missing columns in the new file & add columns from the old file if we need to
for col in initial_worksheet.columns:
    if col not in current_worksheet.columns:
        current_worksheet = current_worksheet.join(initial_worksheet[col])

# parse the region file to dataframe
df = pd.read_excel("src/corona_regions.xlsx")
region_name = df["region"].tolist()
start_points = df["start"].tolist()
end_points = df["end"].tolist()
for i in range(len(region_name)):
    region_name[i] += "(" + str(start_points[i]) + "-" + str(end_points[i]) + ")"
    # create_new columns with regions name and range
    current_worksheet[region_name[i]] = ""
# parse the fasta file.
fasta_sequences = SeqIO.parse(open("src/fasta_aligned_project.fasta"), 'fasta')
# if we move all over the line we will want to stop it and not waste time.
detect_counter = 0
# create the reference sequence - the first one in the fasta file
ref_sequence = str(next(fasta_sequences).seq)
# run over every sequence and if in the fasta file
for fasta in fasta_sequences:
    id, sequence = int(fasta.id), str(fasta.seq)
    # get index of specific seq's id
    id_idx = current_worksheet[current_worksheet["full_sequence_new_sticker_number"] == id].index
    if len(id_idx) > 0:
        detect_counter += 1
        # calculate the number of and the percentage of it in the whole sequence
        n_percentage = sequence.count('n') / len(sequence) * 100
        # we don't want to use sequences which contain 50% of n bases
        if n_percentage > 50:
            continue
        for start, end , col_name in zip(start_points, end_points , region_name):
            index = start - 1
            mutation_counter = 0
            while index != end - 1:
                # we don't compare 'n' at all
                if sequence[index] == 'n':
                    index += 1
                    continue
                # if we have mutation - mismatch or something else
                if sequence[index] != ref_sequence[index]:
                    mutation_counter += 1
                index += 1
            # add sequence to specific index in the fasta column
            current_worksheet.at[id_idx[0], col_name] = mutation_counter
    # don't waste time on empty checks
    if detect_counter == len(current_worksheet.index):
        break

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter("src/output/output.xlsx", engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object. Turn off the default
# header and index and skip one row to allow us to insert a user defined
# header.
current_worksheet.to_excel(writer, sheet_name='Sheet1', startrow=1, header=False, index=False)
# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Get the dimensions of the dataframe.
(max_row, max_col) = current_worksheet.shape

# Create a list of column headers, to use in add_table().
column_settings = []
for header in current_worksheet.columns:
    column_settings.append({'header': header})

# Add the table.
worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

# Make the columns wider for clarity.
worksheet.set_column(0, max_col - 1, 12)

# Close the Pandas Excel writer and output the Excel file.
writer.save()
