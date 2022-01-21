from tkinter import *
from tkinter import filedialog
import textract
import tkinter as tk
from spacy.tokens import Doc
from spacy.training import Example
import spacy
import os

#volledige training data lijst in plaats van laatste element
labelList = ['Naam', 'Geboortedatum', 'Plaats', 'Werkgever', 'Periode', 'Opleidingen', 'Functie', "Verantwoordelijkheden"]
def print_doc_entities(_doc: Doc):
    #print alle entities die een label hebben
    if _doc.ents:
        for _ent in _doc.ents:
            print(f"     {_ent.text} {_ent.label_}")
    else:
        print("     NONE")

def openFile():
    #open een directory. Alle bestanden met .docx of .pdf kunnen worden geopend
    tf = filedialog.askopenfilename(
        initialdir="C:/Users/MainFrame/Desktop/",
        title="Open Text file",
        filetypes=(("Text Files", "*.docx *.pdf"),)
        )
    #decode naar text bestand
    txt = textract.process(tf)
    txt = txt.decode("utf-8")
    #voeg tekst toe aan GUI
    if txtarea == "":
        txtarea.insert(END, txt)
    else:
        #verwijder al bestaande tekst en vervang dit met nieuwe tekst
        txtarea.delete("1.0", "end-1c")
        txtarea.insert(END, txt)
#GUI code toevoegen
ws = Tk()
ws.title("NER learning tool")
#afmetingen GUI scherm
ws.geometry("450x500")
#achtergrondkleur
ws['bg']='#fb0'
#tekstveld
txtarea = Text(ws, width=40, height=20)
txtarea.pack(pady=20)
#lijst met alle labels in de GUI
variable = StringVar(ws)
variable.set(labelList[0])#default value
w = OptionMenu(ws, variable, *labelList)
w.pack()

global TRAIN_DATA
TRAIN_DATA = []
global raw_training_data
global entities
entities = []

def OnButton():
    #ontvang alle tekst in het tekstvak
    raw_training_data = txtarea.get("1.0","end-1c")
    #huidig geselecteerde label
    label = variable.get()
    #geselecteerde tekst
    print("selected text: '%s'" % txtarea.get(tk.SEL_FIRST, tk.SEL_LAST))
    #positie van het begin van de geselecteerde tekst
    locationStart = txtarea.count("1.0", "sel.first")
    if type(locationStart) is tuple:
        locationStart = list(locationStart)
        locationStart = locationStart[0]
        print(locationStart)
    else:
        locationStart = 0
    #positie van het einde van de geselecteerde tekst
    locationEnd = txtarea.count("1.0", "sel.last")
    if type(locationEnd) is tuple:
        locationEnd = list(locationEnd)
        locationEnd = locationEnd[0]
        print(locationEnd)
    #voeg de waardes toe in het format die vereist is door spacy
    entities.append(tuple((locationStart, locationEnd, label)))
    print(entities)
    TRAIN_DATA.append(tuple((raw_training_data, entities)))
def updateModel():
    #de tekst in de GUI is de training data
    raw_training_data = txtarea.get("1.0","end-1c")
    #laad het model in spacy
    nlp = spacy.load(os.getcwd())
    print(nlp)
    #voeg de name entity recognition toe aan het model indien nodig
    if 'ner' not in nlp.pipe_names:
        print("ner not in pipe")
        ner = nlp.add_pipe('ner')
    else:
        print("ner in pipe")
        ner = nlp.get_pipe('ner')
    #https://spacy.io/usage/processing-pipelines kijk hier voor alle verschillende pipeline componenten
    print(nlp.pipe_names)
    #voeg alle custom labels toe aan het model
    for label in range(len(labelList)):
        ner.add_label(labelList[label])
    print(f"\nResult BEFORE training:")
    #ga door met het trainen van het huidige model
    nlp.resume_training()
    #run het model over de huidige tekst
    doc = nlp(raw_training_data)
    #print alle entiteiten in de doc container. https://spacy.io/api/doc
    print_doc_entities(doc)
    optimizer = nlp.create_optimizer()
    print("updating model...")
    #train het meerdere keren per training.
    for _ in range(25):
        for raw_text, entity_offsets in TRAIN_DATA:
            doc = nlp.make_doc(raw_text)
            #maak een example object. Deze worden gebruikt om het model te trainen
            example = Example.from_dict(doc, {"entities": entity_offsets})
            #update het huidige model met een example, dropoutrate van 0.2
            nlp.update([example], drop=0.2, sgd=optimizer)
    print(f"Result AFTER training:")
    #run het model weer na het trainen
    doc = nlp(raw_training_data)
    print_doc_entities(doc)
    #sla het model op in de huidige folder
    nlp.to_disk(os.getcwd())
#set label knop
button = tk.Button(ws, text="Set label", command=OnButton)
button.pack()
#update knop
updateButton = tk.Button(ws, text='Update model', command=updateModel)
updateButton.pack()
#open file knop
Button(
    ws,
    text="Open File",
    command=openFile
    ).pack(side=RIGHT, expand=True, fill=X, padx=20)
#start de GUI
ws.mainloop()


