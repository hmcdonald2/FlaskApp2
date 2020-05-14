from flask import Flask, render_template, url_for, request, session
# note session allows to use cookie
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date
import time
import csv
import docx
from docx import Document


app = Flask(__name__)
app.static_folder = 'static'
# mlist creates blank list to be populated from matterlist.py file
mlist = []
# namelist creates list for drop down menu from mlist
namelist = mlist
# sets the checkbox options
attendtype = ['Call In', 'Call Out', 'Conference', 'Notes', 'Research']
attendupon = ['Client', 'Other Side', 'Counsel', 'Other']
fieldnamestrings = ["mname", "text", "type", "upon", "now"]
# sets the timestamp
now = [time.ctime()]
# todo option to change the time and date to a different time and date
# note this next list effectively sets the orders of the csv rows when saved to csv file
fieldnameslist = fieldnamestrings + now + attendupon + attendtype

# the main page - currently provides an input to create file notes
# refers to the index.html page
@app.route('/', methods=['POST', 'GET'])
def index():
    # remember to send variables as args below (eg. namelist = namelist) where namelist is a variable set in this file.
    return render_template('index.html', namelist = namelist, attendtype = attendtype, attendupon = attendupon)

# takes the information from the index.html page form
# writes the file notes into a csv file
# writes the to do items into a word document
# gives a preview of the file note and to do items that have been saved
# saves the file note text to a PDF in the relevant matter folder
@app.route('/seefn', methods=['GET', 'POST'])
def seefn():
    # writes the text of the file note and the other data to an already existing spreadsheet
    with open("filenotes.csv","a") as file:
        xax = ""
        for el in attendtype:
            if request.form.get(el) == "on":
                xax += el + ", "
        xab = ""
        for el in attendupon:
            if request.form.get(el) == "on":
                xab += el + ", "
        writer = csv.DictWriter(file, fieldnames=fieldnameslist)
        writer.writerow({"text": request.form.get("text"), "mname": request.form.get("mname"), "type": xax, "upon": xab, "now": now[0]})
    # adds the text of the to do list item to a new line in an already existing csv file
    with open("todoexcel.csv","a") as file:
        writer = csv.DictWriter(file,fieldnames=["mname", "To do item", "Date Added"])
        writer.writerow({"mname": request.form.get("mname"), "To do item": request.form.get("tdo"), "Date Added": now[0]})
    # also appends the text of the to do list item to an already existing word doc
    try:
        tdo = request.form.get("tdo")
        doc = docx.Document('todolist.docx')
        doc.add_paragraph(f'{tdo}\n(item added: {now}')
        doc.save('todolist.docx')
    except:
        print("word doc failed")
    # also creates a PDF time stamped file note of the file note
    try:
        from fpdf import FPDF
        import os
        import os.path
        # obtain matter name
        m = request.form.get("mname")
        mname = m.upper()
        # obtain date
        date = now
        # obtain attendance type
        xax = ""
        for el in attendtype:
            if request.form.get(el) == "on":
                xax += el + ", "
        tp = xax.rstrip(", ")
        type = tp.upper()
        # obtain attendance with
        xab = ""
        for el in attendupon:
            if request.form.get(el) == "on":
                xab += el + ", "
        xay = xab.rstrip(" ,")
        attendon = xay.upper()
        # obtain text of file note
        notes = request.form.get("text")
        # create the pdf using the now set variables
        pdf = FPDF()
        pdf.add_page()
        pdf.set_auto_page_break(True, 2)
        pdf.set_font("Arial", size=20)
        pdf.cell(200, 10, txt="FILE NOTE", ln=1, align='C')
        pdf.set_font("Arial", size=12)
        pdf.cell(200, 10, txt=f"Matter name:    {mname}", ln=1, align='L')
        pdf.cell(200, 10, txt=f"Date:             {date}", ln=1, align='l')
        pdf.cell(200, 10, txt=f"Attendance Type: {type}", ln=1, align='l')
        pdf.cell(200, 10, txt=f"Attendance Upon: {attendon}", ln = 1, align = 'l')
        pdf.cell(200, 10, txt="Author: Harry McDonald", ln=1)
        pdf.cell(200, 10, txt="NOTES:", ln=1)
        pdf.multi_cell(200, 10, txt=notes, align='l')
        pdf.ln()
        # save the pdf under the right file name
        for k in range(1, 400):
            if not os.path.exists(f"File Note {mname} {date}.pdf"):
                pdf.output(f'File Note {mname} {date}.pdf')
            elif not os.path.exists(f"File Note {mname} {date}({k}).pdf"):
                pdf.output(f"File Note {mname} {date}({k}).pdf")
                break
        # todo use os.mkdir or shutil move to get file notes into folders sorted by matter
    except:
        print(f"Failed to save the file note as a PDF")
    # following code creates the preview that shows up temporarily on the confirmed page to show what was saved
    txt1 = request.form.get("text")
    txt = txt1.splitlines()
    mtr = request.form.get("mname")
    tdo1 = request.form.get("tdo")
    tdo = tdo1.splitlines()
    return render_template("seefn.html", txt = txt, mtr = mtr, tdo = tdo)




# store the matter list in a .py file
# create functions to add to the list in a permanent way

def createmlist():
    # add the contents of matterlist.py to variable list mlist
    with open ('matterlist.py', 'r') as f:
        filecontents = f.readlines()

        for line in filecontents:
            current_place = line[:-1]
            mlist.append(current_place)
        f.close()

def addmatter(newmatter):
    # function to add new matters to matterlist.py
    with open ('matterlist.py', 'a') as f:
        f.writelines('\n' + newmatter)
        f.close()

def sortmatterlist():
    # function to sort matterlist.py alphabetically
    # be careful because it deletes the whole file and rewrites it
    # also capitalises the entries (if not previously capitalised)
    newlist = []
    # takes the current matterlist.py document, reads it and sorts it
    with open('matterlist.py', 'r') as r:
        for line in r:
            newlist.append(line)
        newlist[-1] += "\n"
        newlist.sort()
    # takes each item read from the document and capitalises it
    newlist = [item.capitalize() for item in newlist]
    for item in newlist:
        item.capitalize()
    # writes (writes over!) the old matterlist.py file with the now sorted and capitalised matter list
    with open('matterlist.py', 'w') as t:
        for line in newlist:
            t.writelines(line)

# run functions in background at start up to create an updated matter list and sort that matter list
createmlist()
sortmatterlist()



if __name__ == "__main__":
    app.run(debug=True)
