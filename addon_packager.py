import wx
from zipfile import ZipFile
import os
import subprocess
from threading import Thread


path = None
if os.path.exists(os.path.join(os.getenv("appdata"), "nvda", "addons")):
	path = os.path.join(os.getenv("appdata"), "nvda", "addons")


class BasePanel(wx.Panel):
	def __init__(self, parent):
		super().__init__(parent)
	def postInit(self):
		self.Parent.sizer.Add(self)
		self.Layout()
		self.Parent.sizer.Fit(self)
		self.SetFocus()






class MainPanel(BasePanel):
	def __init__(self, parent):
		super().__init__(parent)
		wx.StaticText(self, -1, "مسار مجلد الإضافات: ")
		self.pathBox = wx.TextCtrl(self, -1, style=wx.TE_MULTILINE + wx.TE_READONLY + wx.HSCROLL)
		self.browseButton = wx.Button(self, -1, "تغيير...")
		if path:
			self.pathBox.Value = path
		packageButton = wx.Button(self, -1, "تحزيم")
		compressButton = wx.Button(self, -1, "ضغط الإضافات...")
		
		self.browseButton.Bind(wx.EVT_BUTTON, self.onBrowse)
		packageButton.Bind(wx.EVT_BUTTON, self.onPackage)
		compressButton.Bind(wx.EVT_BUTTON, self.onCompress)
		self.postInit()
	def onBrowse(self, event):
		global path
		default_path = path or ""
		selection = wx.DirSelector("اختر مجلد الإضافات", default_path, parent=self.Parent)
		if selection:
			path = selection
			self.pathBox.Value = selection
			self.pathBox.SetFocus()
	def onPackage(self, event):
		if not path:
			self.onBrowse(None)
			self.onPackage(event)
			return
		self.Hide()
		PackagePanel(self.Parent)
	def onCompress(self, event):
		if not path:
			self.onBrowse(None)
			self.onCompress(event)
			return

		self.Hide()
		PackagePanel(self.Parent, "compress")




class PackagePanel(BasePanel):
	def __init__(self, parent, option="package"):
		self.option = option
		super().__init__(parent)
		self.Parent.panels[1] = self
		wx.StaticText(self, -1, "حالة التحزيم: ")
		self.txtStatus = wx.TextCtrl(self, -1, style=wx.TE_READONLY + wx.TE_MULTILINE + wx.HSCROLL) # مكان لعرض التقرير
		self.btnOpen = wx.Button(self, -1, "فتح")
		self.btnOpen.Enabled = False
		self.btnOpen.Bind(wx.EVT_BUTTON, self.onOpen)
		self.btnBack = wx.Button(self, wx.ID_BACKWARD, "رجوع") # الزر الخاص بالرجوع إلى الصفحة السابقة
		self.btnBack.Bind(wx.EVT_BUTTON, self.onBack)
		self.btnBack.Enabled = False
		self.progress = wx.Gauge(self, -1, range=100)
		self.postInit()

		if self.option == "package":
			self.dest = wx.DirSelector("اختر مجلد الإخراج", parent=self.Parent)
		elif self.option == "compress":
			self.dest = wx.SaveFileSelector(" ", ".zip", "my nvda addons", parent=self.Parent)
		if self.dest:
			Thread(target=self.package, args=[path, self.dest]).start()
		else:
			self.onBack()




	def package(self, path, dest):
		wx.CallAfter(self.txtStatus.write, "يتم الآن التحضير لتحزيم الإضافات...\n")
		addon_directories = os.listdir(path)
		os.chdir(path)
		addon_directories = [d for d in addon_directories if os.path.exists(os.path.join(d, "manifest.ini"))]
		if not addon_directories:
			wx.MessageBox("لا توجد إضافات في هذا المجلد", "خطأ", parent=self)
			self.onBack()
			return
		wx.CallAfter(self.txtStatus.write, f"تم العثور على {len(addon_directories)} إضافة في المجلد المحدد\n\n")
		progress = 0
		step = 1/len(addon_directories)
		if self.option == "compress":
			archive = ZipFile(dest, "w")
		for count, directory in enumerate(addon_directories, 1):
			os.chdir(directory)
			wx.CallAfter(self.txtStatus.write, f"تحزيم الإضافة رقم {count}: {directory}\n")
			if self.option == "package":
				addon_path = os.path.join(dest, f"{directory}.nvda-addon")
				if os.path.exists(addon_path):
					os.remove(addon_path)
					wx.CallAfter(self.txtStatus.write, "تم العثور على نسخة سابقة للإضافة وتم حذفها\n")
			else:
				addon_path = os.path.join(path, f"{directory}.nvda-addon")
			addon = ZipFile(addon_path, "w")
			for root, dirs, files in os.walk("."):
				for file in files:
					addon.write(os.path.join(root, file))
			addon.close()
			wx.CallAfter(self.txtStatus.write, f"تم تحزيم الإضافة\n")
			if self.option == "compress":
				archive.write(addon_path, os.path.basename(addon_path))
				os.remove(addon_path)
				wx.CallAfter(self.txtStatus.write, f"تم إضافة {directory} إلى الأرشيف\n\n")
			else:
				wx.CallAfter(self.txtStatus.write, f"المسار: {addon_path}\n\n")
			os.chdir("..")
			progress += step
			wx.CallAfter(self.progress.SetValue, int(progress * 100))
		if self.option == "compress":
			wx.CallAfter(self.txtStatus.write, "تم ضغط إضافاتك وحفظها في المسار المحدد. اضغط على فتح للانتقال إلى الملف المضغوط")
			archive.close()
		else:
			wx.CallAfter(self.txtStatus.write, f"تم تحزيم جميع إضافاتك بنجاح. اضغط على فتح للانتقال إلى مجلد الحفظ")
		self.btnBack.Enabled = self.btnOpen.Enabled = True


	def onOpen(self, event):
		subprocess.run(f"explorer {'/select,' if self.option == 'compress' else ''}{self.dest}")
	def onBack(self, event=None):
		self.Hide()
		del self.Parent.panels[1]
		self.Parent.panels[0].Show()
		self.Parent.panels[0].SetFocus()
		self.Destroy()


class Packager(wx.Frame):
	def __init__(self):
		super().__init__(None, title="محزم الإضافات")
		self.Centre()
		self.sizer = wx.BoxSizer(wx.VERTICAL)
		self.SetSizer(self.sizer)
		self.panels = {0: MainPanel(self)}

		self.Show()
	def onChoose(self, event):
		path = wx.DirSelector("اختر مسار مجلد الإضافات", os.path.join(os.getenv("appdata"), "nvda/addons"), parent=self)
		if not path:
			return
		dest = wx.DirSelector("اختر مسار الإخراج", parent=self)

		if not dest:
			wx.MessageBox("لم تقم بتحديد المجلد الذي سيتم فيه وضع الإضافات المحزمة, لذا فسيتم وضعها في مجلد باسم addons في المستندات", "تنبيه", parent=self)
			try:
				dest = os.path.join(os.getenv("userprofile"), "documents\\addons")
				os.mkdir(dest)
			except FileExistsError:
				pass
		Thread(target=self.package, args=(path, dest)).start()

app = wx.App()
Packager()
app.MainLoop()