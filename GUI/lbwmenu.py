import Tkinter
import tkMessageBox

class ListboxWMenu(Tkinter.Listbox):
    
    def __init__(self, parent, *args, **kwargs):
        Tkinter.Listbox.__init__(self, parent, *args, **kwargs)

        self.click_menu = Tkinter.Menu(self, tearoff=0)
        self.click_menu.add_command(label='Delete', command=self.delete_selected)
        
        self.bind('<Button-3>', self.popup)
        self.bind('<Delete>', self.delete_selected)

    def popup(self, event):
        if self.curselection():
            self.click_menu.tk_popup(event.x_root + 35, event.y_root + 10, 0)

    def delete_selected(self, *args):
        if tkMessageBox.askokcancel('Delete', 'Do you really want to delete these entries?'):
            if self.curselection():
                for entry in self.curselection()[::-1]:
                    self.delete(entry)