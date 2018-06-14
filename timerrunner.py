#Boa:Frame:Frame1

import wx
import os
from threading import Timer
import time
import win32api

def delayrun():
    print 'start...'
   

[wxID_FRAME1, wxID_FRAME1BUTTON1, wxID_FRAME1RADIOBUTTON1, 
 wxID_FRAME1RADIOBUTTON2, wxID_FRAME1RADIOBUTTON3, wxID_FRAME1RADIOBUTTON4, 
] = [wx.NewId() for _init_ctrls in range(6)]

class Frame1(wx.Frame):
    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_FRAME1, name='', parent=prnt,
              pos=wx.Point(655, 308), size=wx.Size(230, 250),
              style=wx.DEFAULT_FRAME_STYLE, title=u'List File Convert')
        self.SetClientSize(wx.Size(214, 212))
        self.SetBackgroundColour(wx.Colour(224, 224, 224))

        self.button1 = wx.Button(id=wxID_FRAME1BUTTON1, label=u'Start',
              name='button1', parent=self, pos=wx.Point(64, 128),
              size=wx.Size(88, 32), style=0)
        self.button1.Bind(wx.EVT_BUTTON, self.OnButton1Button,
              id=wxID_FRAME1BUTTON1)

        self.radioButton1 = wx.RadioButton(id=wxID_FRAME1RADIOBUTTON1,
              label=u'Num Lock', name='radioButton1', parent=self,
              pos=wx.Point(56, 16), size=wx.Size(91, 15), style=0)
        self.radioButton1.SetValue(True)

        self.radioButton2 = wx.RadioButton(id=wxID_FRAME1RADIOBUTTON2,
              label=u'Enter', name='radioButton2', parent=self, pos=wx.Point(56,
              40), size=wx.Size(91, 15), style=0)
        self.radioButton2.SetValue(False)

        self.radioButton3 = wx.RadioButton(id=wxID_FRAME1RADIOBUTTON3,
              label=u'Pause', name='radioButton3', parent=self, pos=wx.Point(56,
              64), size=wx.Size(91, 15), style=0)
        self.radioButton3.SetValue(False)

        self.radioButton4 = wx.RadioButton(id=wxID_FRAME1RADIOBUTTON4,
              label=u'Print Screen', name='radioButton4', parent=self,
              pos=wx.Point(56, 88), size=wx.Size(91, 15), style=0)
        self.radioButton4.SetValue(False)

    def __init__(self, parent):
        self._init_ctrls(parent)
      


                
    def OnButton1Button(self, event):
                
        #input key 
        if self.radioButton1.Value:
            in_key = 144
        if self.radioButton2.Value:
            in_key = 13
        if self.radioButton3.Value:
            in_key = 19
        if self.radioButton4.Value:
            in_key = 42                                    

        time_interval = 1

            
        t = Timer(time_interval,delayrun)
        t.start()
        while True:
            time.sleep(170)
            win32api.keybd_event(in_key,0,0,0)

app = wx.App()
frame = Frame1(None)
frame.Show()
app.MainLoop()

