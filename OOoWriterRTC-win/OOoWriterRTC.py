# -*- coding: cp932 -*-

import optparse
import sys,os,platform
import re
import time
import random
import commands
import math


if os.name == 'posix':
    sys.path += ['/usr/lib/python2.6/dist-packages', '/usr/lib/python2.6/dist-packages/rtctree/rtmidl']
elif os.name == 'nt':
    sys.path += ['C:\\Python26\\lib\\site-packages', 'C:\\Python26\\lib\\site-packages\\rtctree\\rtmidl']

import time
import random
import commands
import RTC
import OpenRTM_aist

from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming





import uno
import unohelper
from com.sun.star.awt import XActionListener

from com.sun.star.script.provider import XScriptContext

from com.sun.star.beans import PropertyValue
from com.sun.star.table import TableBorder
from com.sun.star.text import TableColumnSeparator
from com.sun.star.text.HoriOrientation import NONE as HO_NONE


import OOoRTC





#comp_num = random.randint(1,3000)
imp_id = "OOoWriterControl"# + str(comp_num)







ooowritercontrol_spec = ["implementation_id", imp_id,
                  "type_name",         imp_id,
                  "description",       "Openoffice Writer Component",
                  "version",           "0.1",
                  "vendor",            "Miyamoto Nobuhiko",
                  "category",          "example",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  ""]




##
# OpenOffice Writerを操作するためのRTCのクラス
##

class OOoWriterControl(OpenRTM_aist.DataFlowComponentBase):
  def __init__(self, manager):
    OpenRTM_aist.DataFlowComponentBase.__init__(self, manager)
    
    
    self._d_m_word = RTC.TimedString(RTC.Time(0,0),0)
    self._m_wordIn = OpenRTM_aist.InPort("word", self._d_m_word)

    self._d_m_fontSize = RTC.TimedFloat(RTC.Time(0,0),0)
    self._m_fontSizeIn = OpenRTM_aist.InPort("fontSize", self._d_m_fontSize)

    self._d_m_wsCharacter = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsCharacterIn = OpenRTM_aist.InPort("wsCharacter", self._d_m_wsCharacter)

    self._d_m_wsWord = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsWordIn = OpenRTM_aist.InPort("wsWord", self._d_m_wsWord)

    self._d_m_wsLine = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsLineIn = OpenRTM_aist.InPort("wsLine", self._d_m_wsLine)

    self._d_m_wsParagraph = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsParagraphIn = OpenRTM_aist.InPort("wsParagraph", self._d_m_wsParagraph)

    self._d_m_wsWindow = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsWindowIn = OpenRTM_aist.InPort("wsWindow", self._d_m_wsWindow)

    self._d_m_wsScreen = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsScreenIn = OpenRTM_aist.InPort("wsScreen", self._d_m_wsScreen)

    self._d_m_color = RTC.TimedRGBColour(RTC.Time(0,0),RTC.RGBColour(0,0,0))
    self._m_colorIn = OpenRTM_aist.InPort("color", self._d_m_color)

    self._d_m_selWord = RTC.TimedString(RTC.Time(0,0),0)
    self._m_selWordOut = OpenRTM_aist.OutPort("selWord", self._d_m_selWord)

    self._d_m_copyWord = RTC.TimedString(RTC.Time(0,0),0)
    self._m_copyWordOut = OpenRTM_aist.OutPort("copyWord", self._d_m_copyWord)
    

    try:
      self.writer = OOoWriter()
    except NotOOoWtiterException:
      return

    self.fontSize = 10
    
    return

  ##
  # 実行周期を設定する関数
  ##
  def m_setRate(self, rate):
      m_ec = self.get_owned_contexts()
      m_ec[0].set_rate(rate)

  ##
  # 活性化するための関数
  ## 
  def m_activate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].activate_component(self._objref)

  ##
  # 不活性化するための関数
  ##
  def m_deactivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].deactivate_component(self._objref)

  


  ##
  # 初期化処理用コールバック関数
  ##
  def onInitialize(self):
    
    OOoRTC.writer_comp = self

    self.addInPort("word",self._m_wordIn)
    self.addInPort("fontSize",self._m_fontSizeIn)
    self.addInPort("wsCharacter",self._m_wsCharacterIn)
    self.addInPort("wsWord",self._m_wsWordIn)
    self.addInPort("wsLine",self._m_wsLineIn)
    self.addInPort("wsParagraph",self._m_wsParagraphIn)
    self.addInPort("wsWindow",self._m_wsWindowIn)
    self.addInPort("wsScreen",self._m_wsScreenIn)
    self.addInPort("color",self._m_colorIn)
    self.addOutPort("selWord",self._m_selWordOut)
    self.addOutPort("copyWord",self._m_copyWordOut)

    
    
    
    return RTC.RTC_OK

  ##
  # 文字書き込みの関数
  ##

  def SetWord(self, m_str):
      cursor = self.writer.document.getCurrentController().getViewCursor()

      
      cursor.CharHeight = self.fontSize
      cursor.CharHeightAsian = self.fontSize
       
      cursor.setString(m_str)
      

      cursor.goRight(len(m_str),False)

      cursor.collapseToEnd()

  ##
  # カーソル位置の文字取得の関数
  ##

  def GetWord(self):
      cursor = self.writer.document.getCurrentController().getViewCursor()
       
      return str(cursor.getString())

  
      
      

  ##
  # 文字数移動する関数
  ##
  def MoveCharacter(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      if diff > 0:
          cursor.goRight(diff,False)
          cursor.collapseToEnd()
      else:
          cursor.goLeft(-diff,False)
          cursor.collapseToStart()
          
  ##
  # 単語数移動する関数
  ##
  def MoveWord(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      for i in range(0, diff):
          if diff > 0:
              cursor.gotoNextWord(False)
              cursor.collapseToEnd()
          else:
              cursor.gotoPreviousWord(False)
              cursor.collapseToStart()

  ##
  # 行数移動する関数
  ##
  def MoveLine(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      if diff > 0:
          cursor.goDown(diff,False)
          cursor.collapseToEnd()
      else:
          cursor.goUp(-diff,False)
          cursor.collapseToStart()

  ##
  # 段落数移動する関数
  ##
  def MoveParagraph(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      for i in range(0, diff):
          if diff > 0:
              cursor.gotoNextParagraph(False)
              cursor.collapseToEnd()
          else:
              cursor.gotoPreviousParagraph(False)
              cursor.collapseToStart()


  

  ##
  # 周期処理用コールバック関数
  ##
  
  def onExecute(self, ec_id):
    if self._m_wordIn.isNew():
        data = self._m_wordIn.read()
        self.SetWord(data.data)

    if self._m_fontSizeIn.isNew():
        data = self._m_fontSizeIn.read()
        self.fontSize = data.data

    if self._m_wsCharacterIn.isNew():
        data = self._m_wsCharacterIn.read()
        self.MoveCharacter(data.data)

    if self._m_wsWordIn.isNew():
        data = self._m_wsWordIn.read()
        self.MoveWord(data.data)

    if self._m_wsLineIn.isNew():
        data = self._m_wsLineIn.read()
        self.MoveLine(data.data)

    if self._m_wsParagraphIn.isNew():
        data = self._m_wsParagraphIn.read()
        self.MoveParagraph(data.data)

    if self._m_wsWindowIn.isNew():
        data = self._m_wsWindowIn.read()
        pass


    if self._m_wsScreenIn.isNew():
        data = self._m_wsScreenIn.read()
        pass

    
    self._d_m_selWord.data = str(self.GetWord())
    self._m_selWordOut.write()
        

    return RTC.RTC_OK

  ##
  # 終了処理用コールバック関数
  ##
  
  def on_shutdown(self, ec_id):
      OOoRTC.writer_comp = None
      return RTC.RTC_OK



##
# コンポーネントを活性化してWriterの操作を開始する関数
##

def Start():
    
    if OOoRTC.writer_comp:
        OOoRTC.writer_comp.m_activate()

##
# コンポーネントを不活性化してWriterの操作を終了する関数
##

def Stop():
    
    if OOoRTC.writer_comp:
        OOoRTC.writer_comp.m_deactivate()


##
# コンポーネントの実行周期を設定する関数
##

def Set_Rate():
    pass
    """if OOoRTC.writer_comp:
      try:
        writer = OOoDraw()
      except NotOOoWriterException:
          return

      oWriterPages = draw.drawpages
      for i in range(0, oDrawPages.Count):
        oDrawPage = oDrawPages.getByIndex(i)
        forms = oDrawPage.getForms()
        for j in range(0, forms.Count):
          st_control = oDrawPage.getForms().getByIndex(j).getByName('Rate')
          if st_control:
            try:
              text = float(st_control.Text)
            except:
               return
              
            OOoRTC.draw_comp.m_setRate(text)"""
      
      

      
        
        
      
      

      



      
  



##
#RTCをマネージャに登録する関数
##
def OOoWriterControlInit(manager):
  profile = OpenRTM_aist.Properties(defaults_str=ooowritercontrol_spec)
  manager.registerFactory(profile,
                          OOoWriterControl,
                          OpenRTM_aist.Delete)


def MyModuleInit(manager):
  manager._factory.unregisterObject(imp_id)
  OOoWriterControlInit(manager)

  
  comp = manager.createComponent(imp_id)






          

##
# RTC起動の関数
##

def createOOoWriterComp():
                        
    
    if OOoRTC.mgr == None:
      OOoRTC.mgr = OpenRTM_aist.Manager.init(['OOoWriter.py'])
      OOoRTC.mgr.setModuleInitProc(MyModuleInit)
      OOoRTC.mgr.activateManager()
      OOoRTC.mgr.runManager(True)
    else:
      MyModuleInit(OOoRTC.mgr)
      
          

    try:
      writer = OOoWriter()
    except NotOOoWriterException:
      return

    
    MyMsgBox('',u'RTCを起動しました')


    
    
    return None




##
# メッセージボックス表示の関数
# title：ウインドウのタイトル
# message：表示する文章
##

def MyMsgBox(title, message):
    try:
        m_bridge = Bridge()
    except:
        return
    m_bridge.run_infodialog(title, message)


##
# OpenOfficeを操作するためのクラス
##

class Bridge(object):
  def __init__(self):
    self._desktop = XSCRIPTCONTEXT.getDesktop()
    self._document = XSCRIPTCONTEXT.getDocument()
    self._frame = self._desktop.CurrentFrame
    self._window = self._frame.ContainerWindow
    self._toolkit = self._window.Toolkit
  def run_infodialog(self, title='', message=''):
    msgbox = self._toolkit.createMessageBox(self._window,uno.createUnoStruct('com.sun.star.awt.Rectangle'),'infobox',1,title,message)
    msgbox.execute()
    msgbox.dispose()





##
# OpenOffice Writerを操作するためのクラス
##

class OOoWriter(Bridge):
  def __init__(self):
    Bridge.__init__(self)
    if not self._document.supportsService('com.sun.star.text.TextDocument'):
      self.run_errordialog(title='エラー', message='このマクロはOpenOffice.org Writerの中で実行してください')
      raise NotOOoWriterException()
    self.__current_controller = self._document.CurrentController
    
  @property
  def document(self): return self._document
  


    


g_exportedScripts = (createOOoWriterComp, Start, Stop, Set_Rate)
