"""Native Windows entry point, launched by the Desktop shortcut."""
import os
import queue
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from report_export import prepare, export_reports

class App:
 def __init__(self,root):
  self.root=root;self.files=[None,None];self.events=queue.Queue();self.busy=False
  root.title('ԶՄՄԿ · Excel պլանների համեմատություն');root.geometry('820x610');root.minsize(700,560);root.configure(bg='#f1f5f8')
  icon=Path(__file__).with_name('icon.ico')
  if icon.exists():
   try:root.iconbitmap(str(icon))
   except tk.TclError:pass
  style=ttk.Style();style.theme_use('clam');style.configure('TButton',font=('Segoe UI',11),padding=12);style.configure('Accent.TButton',background='#008c87',foreground='white');style.map('Accent.TButton',background=[('active','#007975')])
  frame=tk.Frame(root,bg='#f1f5f8',padx=35,pady=28);frame.pack(fill='both',expand=True)
  tk.Label(frame,text='ԶՄՄԿ / 2026',bg='#f1f5f8',fg='#008c87',font=('Segoe UI',11,'bold')).pack(anchor='w')
  tk.Label(frame,text='Երկու Excel → համեմատական հաշվետվություն',bg='#f1f5f8',fg='#173249',font=('Segoe UI',19,'bold'),wraplength=735,justify='left').pack(anchor='w',pady=(8,20))
  self.buttons=[];self.labels=[]
  for i,title in enumerate(['1. Ընտրել բազային պլանը','2. Ընտրել թարմացված պլանը']):
   box=tk.Frame(frame,bg='white',padx=15,pady=12);box.pack(fill='x',pady=6)
   button=ttk.Button(box,text=title,command=lambda n=i:self.choose(n));button.pack(anchor='w');self.buttons.append(button)
   label=tk.Label(box,text='Ֆայլը ընտրված չէ',bg='white',fg='#64798c',font=('Segoe UI',10),wraplength=690,justify='left');label.pack(anchor='w',pady=(6,0));self.labels.append(label)
  tk.Label(frame,text='Ստեղծվում են .xlsx և .pdf ֆայլեր՝ նույն գրաֆիկներով։',bg='#f1f5f8',fg='#64798c',font=('Segoe UI',11)).pack(anchor='w',pady=(14,8))
  self.run=ttk.Button(frame,text='Ստեղծել Excel և PDF հաշվետվություններ',style='Accent.TButton',command=self.generate);self.run.pack(anchor='w',pady=8)
  self.progress=ttk.Progressbar(frame,mode='indeterminate');self.progress.pack(fill='x',pady=8)
  self.status=tk.Label(frame,text='Ընտրեք երկու տարբերակները։',bg='#f1f5f8',fg='#64798c',font=('Segoe UI',10),wraplength=720,justify='left');self.status.pack(anchor='w')
  self.output=None;self.open=ttk.Button(frame,text='Բացել արդյունքների պանակը',command=self.open_output,state='disabled');self.open.pack(anchor='w',pady=8)
  root.protocol('WM_DELETE_WINDOW',self.close);root.after(150,self.poll)
 def close(self):
  if self.busy:messagebox.showinfo('Մշակվում է','Սպասեք հաշվետվությունների ստեղծման ավարտին։');return
  self.root.destroy()
 def choose(self,index):
  path=filedialog.askopenfilename(title='Ընտրեք Excel պլանը',filetypes=[('Excel','*.xlsx')])
  if path:self.files[index]=path;self.labels[index].configure(text=Path(path).name)
 def generate(self):
  if not all(self.files):messagebox.showwarning('Ընտրեք ֆայլերը','Անհրաժեշտ է ընտրել երկու Excel ֆայլ։');return
  if Path(self.files[0]).resolve()==Path(self.files[1]).resolve():messagebox.showwarning('Նույն ֆայլը','Ընտրեք երկու տարբեր Excel ֆայլեր։');return
  initial=Path.home()/'Downloads'
  folder=filedialog.askdirectory(title='Որտեղ պահպանել Excel և PDF արդյունքները',initialdir=str(initial if initial.exists() else Path.home()))
  if not folder:return
  self.busy=True;self.run.configure(state='disabled');self.open.configure(state='disabled')
  for b in self.buttons:b.configure(state='disabled')
  self.progress.start();self.status.configure(text='Ֆայլերը մշակվում են, ստեղծվում են գրաֆիկները…')
  threading.Thread(target=self.worker,args=(list(self.files),folder),daemon=True).start()
 def worker(self,files,folder):
  target=None
  try:
   result=prepare(*files)
   target=Path(folder)/('ZCMC_'+datetime.now().strftime('%Y%m%d_%H%M%S_%f'));target.mkdir()
   export_reports(result,target/'ZCMC_comparison.xlsx',target/'ZCMC_comparison.pdf')
   self.events.put(('ok',target))
  except Exception as error:
   if target and target.exists():
    # Remove only known incomplete outputs in the unique directory created above.
    for name in ['ZCMC_comparison.xlsx','ZCMC_comparison.pdf']:
     p=target/name
     if p.exists():p.unlink()
    try:target.rmdir()
    except OSError:pass
   self.events.put(('error',str(error)))
 def poll(self):
  try:
   kind,data=self.events.get_nowait();self.busy=False;self.progress.stop();self.run.configure(state='normal')
   for b in self.buttons:b.configure(state='normal')
   if kind=='ok':
    self.output=data;self.open.configure(state='normal');self.status.configure(text='Պատրաստ է՝ '+str(data),fg='#008c87');messagebox.showinfo('Պատրաստ է','Excel և PDF հաշվետվությունները ստեղծված են։\n\n'+str(data))
   else:self.status.configure(text='Ստեղծումը չհաջողվեց։',fg='#a32828');messagebox.showerror('Սխալ',data)
  except queue.Empty:pass
  self.root.after(150,self.poll)
 def open_output(self):
  if self.output:os.startfile(str(self.output))

if __name__=='__main__':
 root=tk.Tk();App(root);root.mainloop()
