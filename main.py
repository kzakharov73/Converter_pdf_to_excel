from tkinter import Label, Button, filedialog
import aspose.pdf as ap
import tkinter as tk


def select_pdf_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    inp.delete(0, 'end')
    inp.insert(0, file_path)


def convert_pdf_to_xlsx():
    pdf_path = inp.get()

    xlsx_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    out.delete(0, 'end')
    out.insert(0, xlsx_path)

    if pdf_path and xlsx_path:

        # Open PDF document
        document = ap.Document(pdf_path)

        save_option = ap.ExcelSaveOptions()
        save_option.format = ap.ExcelSaveOptions.ExcelFormat.XLSX
        save_option.minimize_the_number_of_worksheets = True
        # Save the file
        document.save(xlsx_path, save_option)

        result_label.config(text='Conversion completed successfully!')
    else:
        result_label.config(text='Please select file')


# GUI setup
root = tk.Tk()
root.title('PDF to XLSX Converter')
root.geometry('750x100')
root['bg'] = '#696969'
inp = tk.Entry(root, width=90)
inp.grid(row=1, column=0)

select_pdf_button = Button(root, text='Select PDF', command=select_pdf_file, bg='#FF0000', width=15,
                           font=('Arial', 10, 'bold'))
select_pdf_button.grid(row=1, column=1)

out = tk.Entry(root, width=90)
out.grid(row=2, column=0)
convert_button = Button(root, text='Convert to XLSX', command=convert_pdf_to_xlsx, bg='#32CD32', width=15,
                        font=('Arial', 10, 'bold'))
convert_button.grid(row=2, column=1)

result_label = Label(root, text='', font=('Arial', 14), fg='#4B0082', bg='#696969')
result_label.grid(column=0, row=4)

info_label = Label(root, text='PDF ==>> XLSX', font=('Arial', 16, 'bold'), fg='#800080', bg='#696969')
info_label.grid(row=4, column=1)
root.mainloop()
