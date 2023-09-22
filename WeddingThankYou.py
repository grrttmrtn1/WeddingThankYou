import tkinter as tk
import tkinter.font as tkFont
from tkinter import filedialog
import pandas as pd
import openai

class App:
    def __init__(self, root):
        #setting title
        root.title("Wedding Thank You")
        #setting window size
        width=230
        height=120
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        root.resizable(width=False, height=False)

        GButton_688=tk.Button(root)
        GButton_688["bg"] = "#e9e9ed"
        ft = tkFont.Font(family='Times',size=10)
        GButton_688["font"] = ft
        GButton_688["fg"] = "#000000"
        GButton_688["justify"] = "center"
        GButton_688["text"] = "Open Guest List"
        GButton_688.place(x=10,y=10,width=98,height=30)
        GButton_688["command"] = self.openFile

        self.GLineEdit_720=tk.Entry(root)
        self.GLineEdit_720["borderwidth"] = "1px"
        ft = tkFont.Font(family='Times',size=10)
        self.GLineEdit_720["font"] = ft
        self.GLineEdit_720["fg"] = "#333333"
        self.GLineEdit_720["justify"] = "center"
        self.GLineEdit_720["text"] = "Input API Key"
        self.GLineEdit_720.insert(0, "Input API Key")
        self.GLineEdit_720.place(x=120,y=10,width=98,height=30)

        self.GLabel_559=tk.Label(root)
        ft = tkFont.Font(family='Times',size=10)
        self.GLabel_559["font"] = ft
        self.GLabel_559["fg"] = "#333333"
        self.GLabel_559["justify"] = "center"
        self.GLabel_559["text"] = "Status: pending..."
        self.GLabel_559.place(x=50,y=50,width=140,height=25)

        GButton_38=tk.Button(root)
        GButton_38["bg"] = "#e9e9ed"
        ft = tkFont.Font(family='Times',size=10)
        GButton_38["font"] = ft
        GButton_38["fg"] = "#000000"
        GButton_38["justify"] = "center"
        GButton_38["text"] = "Submit"
        GButton_38.place(x=80,y=80,width=70,height=25)
        GButton_38["command"] = self.submit


    def openFile(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel file",".csv")])
        self.df = pd.read_csv(self.file_path)

    def submit(self):
        self.GLabel_559["text"] = "Status: starting..."
        if self.checkAPIKey():
            for i in range(len(self.df)):
                message = f"Write a {self.df.loc[i, 'Event']} thank you card to {self.df.loc[i, 'Guest']} that has the relationship {self.df.loc[i, 'Relationship']} and gave the gift {self.df.loc[i, 'Present']}"
                if 'TRUE' in str(self.df.loc[i, 'DidNotTalk']):
                    message = message + 'and apologize for being unable to get a chance to speak with them at the wedding'
                print(message) 
                response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[{
                        "role": "user",
                        "content": message
                    }],
                temperature=1,
                max_tokens=2048
                )
                self.df.loc[i,  'Thankyoumessage'] = response['choices'][0]['message']['content']
                print(response['choices'][0]['message']['content'])
                self.saveFile()
        else:
            print('failed')

    def checkAPIKey(self):
        self.GLabel_559["text"] = "Status: Checking API Key..."
        self.apiValue = self.GLineEdit_720.get()
        if ' ' in self.apiValue or 'Input API Key' in self.apiValue:
            self.GLabel_559["text"] = "Status: Check Your API Key..."
            return False
        openai.api_key = self.apiValue
        return True

    def saveFile(self):
        try:
            self.df.to_csv(self.file_path)
            self.GLabel_559["text"] = "Status: Saved to file..."
        except:
            self.GLabel_559["text"] = "Status: Failed to Save..."

    
        

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
