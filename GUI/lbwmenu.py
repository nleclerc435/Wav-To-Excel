import Tkinter
import tkMessageBox
import os
import subprocess

class ListboxWMenu(Tkinter.Listbox):
    
    def __init__(self, parent, *args, **kwargs):
        Tkinter.Listbox.__init__(self, parent, *args, **kwargs)
        self.path = ''

        self.click_menu = Tkinter.Menu(self, tearoff=0)
        self.click_menu.add_command(label='Open File', command=self.open_file)
        self.click_menu.add_command(label='Delete', command=self.delete_selected)
        
        self.bind('<Button-3>', self.popup)
        self.bind('<Delete>', self.delete_selected)
        self.bind('<Double-Button-1>', self.open_file)

    def popup(self, event):
        if self.curselection():
            self.click_menu.tk_popup(event.x_root + 35, event.y_root + 10, 0)

    def delete_selected(self, *args):
        if tkMessageBox.askokcancel('Delete', 'Do you really want to delete these entries?'):
            if self.curselection():
                for entry in self.curselection()[::-1]:
                    self.delete(entry)

    def open_file(self, *args):
        wmp = "C:\\Program Files (x86)\\Windows Media Player\\wmplayer.exe"
        if len(self.curselection()) > 1:
            if tkMessageBox.askokcancel('Open','Do you really want to open multiple files?'):
                items = [self.get(i).split('    Duration: ')[0] for i in self.curselection()]
                for i in items:
                    subprocess.call([wmp, self.path+i])
        else:
            if self.curselection():
                filename = self.get(self.curselection()[0]).split('    Duration: ')[0]
                os.startfile(self.path+filename)