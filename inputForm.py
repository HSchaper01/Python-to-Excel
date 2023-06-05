# importing libraries
import pandas as pd
import tkinter as tk
import customtkinter

## we'll have a customkinter section and a pandas section


################################################################################## BEGINNING OF PANDAS SECTION #######################################################################

rCEDatabase = pd.DataFrame(columns=['policyNumber', 'currentRCE', 'agentEmail', 'currentCoverageA', 'insuredName', 'yearBuilt'])

# We'll use pandas to export user input to Excel 
def save_to_excel():
    global rCEDatabase
# Getting the text from the user entry fields
    policyNumber = policyNumberField.get()
    currentRCE = currentRCEField.get()
    agentEmail = agentEmailField.get()
    currentCoverageA = currentCoverageAField.get()
    insuredName = insuredNameField.get()
    yearBuilt = yearBuiltField.get()

# These will be the columns for our Excel sheet
    formData = {'policyNumber':[policyNumber], 'currentRCE':[currentRCE], 'agentEmail':[agentEmail], 'currentCoverageA':[currentCoverageA], 'insuredName':[insuredName],
            'yearBuilt':[yearBuilt]}

    newFormData = pd.DataFrame(formData)

# This will help us append whatever's submitted to the next available roW
    rCEDatabase = pd.concat([rCEDatabase, newFormData], ignore_index=True)
    rCEDatabaseFilePath = (r"C:\Users\hrsch\OneDrive\Desktop\rCEDatabase.xlsx")
    rCEDatabase.to_excel(rCEDatabaseFilePath, index = False)
    clear_fields()
    confLabel.configure(text="Submitted!\n Bringing you back . . .")
    window.after(3000, lambda: (confLabel.configure(text=""), policyNumberField.focus_set()))

def clear_fields():
    policyNumberField.delete(0, tk.END)
    currentRCEField.delete(0, tk.END)
    agentEmailField.delete(0, tk.END)
    currentCoverageAField.delete(0, tk.END)
    insuredNameField.delete(0, tk.END)
    yearBuiltField.delete(0, tk.END)

def conf_message():
    print("The information has been submitted!")

################################################################################## END OF PANDAS SECTION #############################################################################


################################################################################## CUSTOMKINTER SECTION #######################################################################

# customkinter allows us to set the theme
customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

# window will be set as the root and titled "Replacement Cost Database"
window = customtkinter.CTk()
window.geometry("600x745")
window.title("Replacement Cost Database")

# this will be the display name
label = customtkinter.CTkLabel(window, text = "Replacement Cost Database", font = ('Robotica', 18))
label.pack(padx= 20, pady = 20)

# we'll set the rest of the form as a frame
fieldFrame = customtkinter.CTkFrame(window)
fieldFrame.columnconfigure(0, weight = 1)

# these will be the entry fields

# Starting with Policy Number
policyNumberFieldLabel = customtkinter.CTkLabel(master = fieldFrame, text = "Policy Number", font=('Arial', 22)) #LABEL
policyNumberFieldLabel.grid(pady=10, padx=20, row=0, column = 0, sticky = tk.W)                                  #LABEL
policyNumberField = customtkinter.CTkEntry(fieldFrame, placeholder_text="Policy Number")                         #FIELD
policyNumberField.grid(row=1, column = 0, sticky = tk.W, padx = 20, pady = 2)                                    #FIELD

padding_label1 = customtkinter.CTkLabel(fieldFrame, text="")                                                     #SPACER
padding_label1.grid(pady = 5, row = 2)                                                                           #SPACER

# Current RCE
currentRCEFieldLabel = customtkinter.CTkLabel(master = fieldFrame, text = "Current RCE", font=('Arial', 22))
currentRCEFieldLabel.grid(pady=10, padx=20, row=3, column = 0, sticky = tk.W)
currentRCEField = customtkinter.CTkEntry(fieldFrame, placeholder_text="Current RCE")
currentRCEField.grid(row=4, column = 0, sticky = tk.W, padx = 20, pady = 2)

padding_label2 = customtkinter.CTkLabel(fieldFrame, text="")
padding_label2.grid(pady = 5, row = 5)

# Agent Email
agentEmailFieldLabel = customtkinter.CTkLabel(master = fieldFrame, text = "Agent Email", font=('Arial', 22))
agentEmailFieldLabel.grid(pady=10, padx=20, row=6, column = 0, sticky = tk.W)
agentEmailField = customtkinter.CTkEntry(fieldFrame, placeholder_text="Agent Email")
agentEmailField.grid(row=7, column = 0, sticky = tk.W, padx = 20, pady = 2)

padding_label3 = customtkinter.CTkLabel(fieldFrame, text="")
padding_label3.grid(pady = 5, row = 8)

# Current Coverage
currentCoverageAFieldLabel = customtkinter.CTkLabel(master = fieldFrame, text = "Current Coverage", font=('Arial', 22))
currentCoverageAFieldLabel.grid(pady=10, padx=20, row=9, column = 0, sticky = tk.W)
currentCoverageAField = customtkinter.CTkEntry(fieldFrame, placeholder_text="Current Coverage")
currentCoverageAField.grid(row=10, column = 0, sticky = tk.W, padx = 20, pady = 2)

padding_label4 = customtkinter.CTkLabel(fieldFrame, text="")
padding_label4.grid( pady = 5, row = 11)

# Insured Name 
insuredNameFieldLabel = customtkinter.CTkLabel(master = fieldFrame, text = "Insured Name", font=('Arial', 22))
insuredNameFieldLabel.grid(pady=10, padx=20, row=12, column = 0, sticky = tk.W)
insuredNameField = customtkinter.CTkEntry(fieldFrame, placeholder_text = "Insured Name")
insuredNameField.grid(row=13, column = 0, sticky = tk.W, padx = 20, pady = 2)

padding_label5 = customtkinter.CTkLabel(fieldFrame, text="")
padding_label5.grid(pady = 5, row = 14)

# Year Built 
yearBuiltFieldLabel = customtkinter.CTkLabel(master = fieldFrame, text = "Year Built", font=('Arial', 22))
yearBuiltFieldLabel.grid(pady=10, padx=20, row=15, column = 0, sticky = tk.W)
yearBuiltField = customtkinter.CTkEntry(fieldFrame, placeholder_text = "Year Built")
yearBuiltField.grid(row=16, column = 0, sticky = tk.W, padx = 20, pady = 2)


# Button section 
submitButton = customtkinter.CTkButton(master = fieldFrame, text = "Submit", command = save_to_excel)
submitButton.grid(pady=0, padx = 10, row = 1, column = 2)

# Confirmation lable
confLabel = customtkinter.CTkLabel(master=fieldFrame, text = "", font=('Arial', 22))
confLabel.grid(row = 3, column = 1, sticky = tk.E, padx = 5, pady = 2)

fieldFrame.pack(fill = 'x')

policyNumberField.focus_set()

window.mainloop()

################################################################################## END OF CUSTOMKINTER SECTION #######################################################################












