# -*- coding: utf-8 -*-

##
#
# @file OOoWriterRTC.py

import optparse
import sys,os,platform
import re
import time
import random
import commands
import math

import os.path

from os.path import expanduser
sv = sys.version_info


if os.name == 'posix':
    home = expanduser("~")
    sys.path += [home+'/OOoRTC', home+'/OOoRTC/WriterIDL', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages/rtctree/rtmidl']
elif os.name == 'nt':
    sys.path += ['.\\OOoRTC', '.\\OOoRTC\\WriterIDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages', 'C:\\Python2' + str(sv[1]) + '\\Lib\\site-packages\\OpenRTM_aist\\RTM_IDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages\\rtctree\\rtmidl']
    



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

from com.sun.star.awt.FontWeight import BOLD
from com.sun.star.awt.FontWeight import NORMAL as FW_NORMAL
from com.sun.star.awt.FontSlant import ITALIC
from com.sun.star.awt.FontSlant import NONE as FS_NONE
from com.sun.star.awt.FontUnderline import SINGLE as FU_SINGLE
from com.sun.star.awt.FontUnderline import NONE as FU_NONE
from com.sun.star.awt.FontStrikeout import SINGLE as FST_SINGLE
from com.sun.star.awt.FontStrikeout import NONE as FST_NONE
#from com.sun.star.awt.FontEmphasis import NONE as FE_NONE
#from com.sun.star.awt.FontRelief import ENGRAVED
#from com.sun.star.awt.FontRelief import NONE as FR_NONE
from com.sun.star.awt import XActionListener

from com.sun.star.script.provider import XScriptContext

from com.sun.star.beans import PropertyValue
from com.sun.star.table import TableBorder
from com.sun.star.text import TableColumnSeparator
from com.sun.star.text.HoriOrientation import NONE as HO_NONE


import OOoRTC



import Writer_idl

from omniORB import PortableServer
import Writer, Writer__POA




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
                  "conf.default.fontsize", "16",
                  #"conf.default.fontname", "ＭＳ 明朝",
                  "conf.default.Char_Red", "0",
                  "conf.default.Char_Blue", "0",
                  "conf.default.Char_Green", "0",
                  "conf.default.Italic", "0",
                  "conf.default.Bold", "0",
                  "conf.default.Underline", "0",
                  "conf.default.Shadow", "0",
                  "conf.default.Strikeout", "0",
                  "conf.default.Contoured", "0",
                  "conf.default.Emphasis", "0",
                  "conf.default.Back_Red", "255",
                  "conf.default.Back_Blue", "255",
                  "conf.default.Back_Green", "255",
                  "conf.default.Code", "utf-8",
                  "conf.__widget__.fontsize", "spin",
                  #"conf.__widget__.fontname", "radio",
                  "conf.__widget__.Char_Red", "spin",
                  "conf.__widget__.Char_Blue", "spin",
                  "conf.__widget__.Char_Green", "spin",
                  "conf.__widget__.Italic", "radio",
                  "conf.__widget__.Bold", "radio",
                  "conf.__widget__.Underline", "radio",
                  "conf.__widget__.Shadow", "radio",
                  "conf.__widget__.Strikeout", "radio",
                  "conf.__widget__.Contoured", "radio",
                  "conf.__widget__.Emphasis", "radio",
                  "conf.__widget__.Back_Red", "spin",
                  "conf.__widget__.Back_Blue", "spin",
                  "conf.__widget__.Back_Green", "spin",
                  "conf.__widget__.Code", "radio",
                  "conf.__constraints__.fontsize", "1<=x<=72",
                  #"conf.__constraints__.fontname", "(MS UI Gothic,MS ゴシック,MS Pゴシック,MS 明朝,MS P明朝,HG ゴシック E,HGP ゴシック E,HGS ゴシック E,HG ゴシック M,HGP ゴシック M,HGS ゴシック M,HG 正楷書体-PRO,HG 丸ゴシック M-PRO,HG 教科書体,HGP 教科書体,HGS 教科書体,HG 行書体,HGP 行書体,HGS 行書体,HG 創英プレゼンス EB,HGP 創英プレゼンス EB,HGS 創英プレゼンス EB,HG 創英角ゴシック UB,HGP 創英角ゴシック UB,HGS 創英角ゴシック UB,HG 創英角ポップ体,HGP 創英角ポップ体,HGS 創英角ポップ体,HG 明朝 B,HGP 明朝 B,HGS 明朝 B,HG 明朝 E,HGP 明朝 E,HGS 明朝 E,メイリオ)",
                  "conf.__constraints__.Char_Red", "0<=x<=255",
                  "conf.__constraints__.Char_Blue", "0<=x<=255",
                  "conf.__constraints__.Char_Green", "0<=x<=255",
                  "conf.__constraints__.Italic", "(0,1)",
                  "conf.__constraints__.Bold", "(0,1)",
                  "conf.__constraints__.Underline", "(0,1)",
                  "conf.__constraints__.Shadow", "(0,1)",
                  "conf.__constraints__.Strikeout", "(0,1)",
                  "conf.__constraints__.Contoured", "(0,1)",
                  "conf.__constraints__.Emphasis", "(0,1)",
                  "conf.__constraints__.Back_Red", "0<=x<=255",
                  "conf.__constraints__.Back_Blue", "0<=x<=255",
                  "conf.__constraints__.Back_Green", "0<=x<=255",
                  "conf.__constraints__.Code", "(utf-8,euc_jp,shift_jis)",
    ""
                  ""]



##
# @class mWriter_i
# @brief サービスポートWriter
#
class mWriter_i (Writer__POA.mWriter):

    ##
    # @brief コンストラクタ
    # @param self
    # @param m_comp OOoWriterRTC
    #
    def __init__(self, m_comp):
        self.m_comp = m_comp

    ##
    # @brief 現在のカーソル位置X座標取得(単位はmm)
    # @param self
    # @return カーソル位置X座標
    #
    def oCurrentCursorPositionX(self): 
        x,y = self.m_comp.oCurrentCursorPosition()
        return float(x)
        
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    ##
    # @brief 現在のカーソル位置Y座標取得(単位はmm)
    # @param self
    # @return カーソル位置Y座標
    #
    def oCurrentCursorPositionY(self):
        x,y = self.m_comp.oCurrentCursorPosition()
        return float(y)
        
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)

    ##
    # @brief カーソルをドキュメント先頭に移動
    # @param self
    # @param sel Trueなら移動範囲を選択
    #
    def gotoStart(self, sel):
        self.m_comp.gotoStart(sel)
        return
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    
    ##
    # @brief カーソルをドキュメント最後尾に移動
    # @param self
    # @param sel Trueなら移動範囲を選択
    #
    def gotoEnd(self, sel):
        self.m_comp.gotoEnd(sel)
        return
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    
    ##
    # @brief カーソルを行先頭に移動
    # @param self
    # @param sel Trueなら移動範囲を選択
    #
    def gotoStartOfLine(self, sel):
        self.m_comp.gotoStartOfLine(sel)
        return
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        

    
    ##
    # @brief カーソルを行最後尾に移動
    # @param self
    # @param sel Trueなら移動範囲を選択
    #
    def gotoEndOfLine(self, sel):
        self.m_comp.gotoEndOfLine(sel)
        return
        raise CORBA.NO_IMPLEMENT(0, CORBA.COMPLETED_NO)
        
        



##
# @brief ユニコード文字列をドキュメント上で文字化けしない文字コードで文字列を返す
# @param m_str 変換前の文字列
# @return 変換後の文字列
#
def ResetCoding(m_str):
    if os.name == 'posix':
        return m_str.encode('utf-8')
    elif os.name == 'nt':
        return m_str.encode('cp932')








##
# @class OOoWriterControl
# @brief OpenOffice Writerを操作するためのRTCのクラス
#

class OOoWriterControl(OpenRTM_aist.DataFlowComponentBase):
  def __init__(self, manager):
    OpenRTM_aist.DataFlowComponentBase.__init__(self, manager)
    
    
    self._d_m_word = RTC.TimedString(RTC.Time(0,0),0)
    self._m_wordIn = OpenRTM_aist.InPort("word", self._d_m_word)

    self._d_m_fontSize = RTC.TimedFloat(RTC.Time(0,0),0)
    self._m_fontSizeIn = OpenRTM_aist.InPort("fontSize", self._d_m_fontSize)

    self._d_m_fontName = RTC.TimedString(RTC.Time(0,0),0)
    self._m_fontNameIn = OpenRTM_aist.InPort("fontName", self._d_m_fontName)

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

    self._d_m_Char_color = RTC.TimedRGBColour(RTC.Time(0,0),RTC.RGBColour(0,0,0))
    self._m_Char_colorIn = OpenRTM_aist.InPort("Char_color", self._d_m_Char_color)

    self._d_m_MovementType = RTC.TimedBoolean(RTC.Time(0,0),0)
    self._m_MovementTypeIn = OpenRTM_aist.InPort("MovementType", self._d_m_MovementType)

    self._d_m_Italic = RTC.TimedBoolean(RTC.Time(0,0),0)
    self._m_ItalicIn = OpenRTM_aist.InPort("Italic", self._d_m_Italic)

    self._d_m_Bold = RTC.TimedBoolean(RTC.Time(0,0),0)
    self._m_BoldIn = OpenRTM_aist.InPort("Bold", self._d_m_Bold)

    self._d_m_Underline = RTC.TimedBoolean(RTC.Time(0,0),0)
    self._m_UnderlineIn = OpenRTM_aist.InPort("Underline", self._d_m_Underline)

    self._d_m_Shadow = RTC.TimedBoolean(RTC.Time(0,0),0)
    self._m_ShadowIn = OpenRTM_aist.InPort("Shadow", self._d_m_Shadow)

    self._d_m_Strikeout = RTC.TimedBoolean(RTC.Time(0,0),0)
    self._m_StrikeoutIn = OpenRTM_aist.InPort("Strikeout", self._d_m_Strikeout)

    self._d_m_Contoured = RTC.TimedBoolean(RTC.Time(0,0),0)
    self._m_ContouredIn = OpenRTM_aist.InPort("Contoured", self._d_m_Contoured)

    self._d_m_Emphasis = RTC.TimedBoolean(RTC.Time(0,0),0)
    self._m_EmphasisIn = OpenRTM_aist.InPort("Emphasis", self._d_m_Emphasis)

    self._d_m_Back_color = RTC.TimedRGBColour(RTC.Time(0,0),RTC.RGBColour(0,0,0))
    self._m_Back_colorIn = OpenRTM_aist.InPort("Back_color", self._d_m_Back_color)


    self._d_m_selWord = RTC.TimedString(RTC.Time(0,0),0)
    self._m_selWordOut = OpenRTM_aist.OutPort("selWord", self._d_m_selWord)

    self._d_m_copyWord = RTC.TimedString(RTC.Time(0,0),0)
    self._m_copyWordOut = OpenRTM_aist.OutPort("copyWord", self._d_m_copyWord)


    self._WriterPort = OpenRTM_aist.CorbaPort("Writer")
    self._writer = mWriter_i(self)
    

    try:
      self.writer = OOoWriter()
    except NotOOoWtiterException:
      return

    self.fontSize = 16
    self.fontName = "ＭＳ 明朝"
    self.Bold = False
    self.Italic = False
    self.Char_Red = 0
    self.Char_Green = 0
    self.Char_Blue = 0
    self.MovementType = False

    self.Underline = False
    self.Shadow = False
    self.Strikeout = False
    self.Contoured = False
    self.Emphasis = False

    self.Back_Red = 255
    self.Back_Green = 255
    self.Back_Blue = 255


    self.conf_fontSize = [16]
    self.conf_fontName = ["ＭＳ 明朝"]
    self.conf_Bold = [0]
    self.conf_Italic = [0]
    self.conf_Char_Red = [0]
    self.conf_Char_Green = [0]
    self.conf_Char_Blue = [0]
    self.conf_Code = ["utf-8"]

    self.conf_Underline = [0]
    self.conf_Shadow = [0]
    self.conf_Strikeout = [0]
    self.conf_Contoured = [0]
    self.conf_Emphasis = [0]
    self.conf_Back_Red = [255]
    self.conf_Back_Green = [255]
    self.conf_Back_Blue = [255]

    self.file = None
    
    
    return

  ##
  # @brief 実行周期を設定する関数
  # @param self
  # @param rate：実行周期
  #
  
  def m_setRate(self, rate):
      m_ec = self.get_owned_contexts()
      m_ec[0].set_rate(rate)

  ##
  # @brief 活性化するための関数
  # @param self
  #
  def m_activate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].activate_component(self._objref)

  ##
  # @brief 不活性化するための関数
  # @param self
  #
  def m_deactivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].deactivate_component(self._objref)

  


  ##
  # @brief 初期化処理用コールバック関数
  # @param self
  #
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
    self.addInPort("Char_color",self._m_Char_colorIn)
    self.addInPort("MovementType",self._m_MovementTypeIn)
    self.addInPort("Italic",self._m_ItalicIn)
    self.addInPort("Bold",self._m_BoldIn)
    self.addInPort("Underline",self._m_UnderlineIn)
    self.addInPort("Shadow",self._m_ShadowIn)
    self.addInPort("Strikeout",self._m_StrikeoutIn)
    self.addInPort("Contoured",self._m_ContouredIn)
    self.addInPort("Emphasis",self._m_EmphasisIn)
    self.addInPort("Back_color",self._m_Back_colorIn)
    self.addOutPort("selWord",self._m_selWordOut)
    self.addOutPort("copyWord",self._m_copyWordOut)

    self._WriterPort.registerProvider("writer", "Writer::mWriter", self._writer)
    self.addPort(self._WriterPort)

    self.bindParameter("fontsize", self.conf_fontSize, "16")
    #self.bindParameter("fontname", self.conf_fontName, "ＭＳ 明朝")
    self.bindParameter("Bold", self.conf_Bold, "0")
    self.bindParameter("Italic", self.conf_Italic, "0")
    self.bindParameter("Char_Red", self.conf_Char_Red, "0")
    self.bindParameter("Char_Blue", self.conf_Char_Blue, "0")
    self.bindParameter("Char_Green", self.conf_Char_Green, "0")
    self.bindParameter("Code", self.conf_Code, "utf-8")

    self.bindParameter("Underline", self.conf_Underline, "0")
    self.bindParameter("Shadow", self.conf_Shadow, "0")
    self.bindParameter("Strikeout", self.conf_Strikeout, "0")
    self.bindParameter("Contoured", self.conf_Contoured, "0")
    self.bindParameter("Emphasis", self.conf_Emphasis, "0")

    self.bindParameter("Back_Red", self.conf_Back_Red, "255")
    self.bindParameter("Back_Blue", self.conf_Back_Blue, "255")
    self.bindParameter("Back_Green", self.conf_Back_Green, "255")

    
    
    
    return RTC.RTC_OK

  ##
  # @brief 文字書き込みの関数
  # @param self
  # @param m_str 書き込む文字列
  #

  def SetWord(self, m_str):
      cursor = self.writer.document.getCurrentController().getViewCursor()

      inp_str = OOoRTC.SetCoding(m_str, self.conf_Code[0])
      cursor.setString(inp_str)
      
      cursor.CharHeight = self.fontSize
      cursor.CharHeightAsian = self.fontSize

      

      if self.Bold:
          cursor.CharWeight = BOLD
          cursor.CharWeightAsian = BOLD
      else:
          cursor.CharWeight = FW_NORMAL
          cursor.CharWeightAsian = FW_NORMAL
      if self.Italic:
          cursor.CharPosture = ITALIC
          cursor.CharPostureAsian = ITALIC
      else:
          cursor.CharPosture = FS_NONE
          cursor.CharPostureAsian = FS_NONE

      if self.Underline:
          cursor.CharUnderline = FU_SINGLE
      else:
          cursor.CharPosture = FU_NONE
          

      if self.Strikeout:
          cursor.CharStrikeout = FST_SINGLE
      else:
          cursor.CharStrikeout = FST_NONE

      if self.Emphasis:
          cursor.CharEmphasis = 1
      else:
          cursor.CharEmphasis = 0

      cursor.CharShadowed = self.Shadow

      cursor.CharContoured = self.Contoured
          

        

      #cursor.CharStyleName = self.fontName

      cursor.CharColor = OOoRTC.RGB(self.Char_Red,self.Char_Green,self.Char_Blue)
      cursor.CharBackColor = OOoRTC.RGB(self.Back_Red,self.Back_Green,self.Back_Blue)

      cursor.goRight(len(inp_str),False)

      cursor.collapseToEnd()

  ##
  # @brief カーソル位置の文字取得の関数
  # @param self
  # @return カーソル位置の文字列
  #

  def GetWord(self):
      cursor = self.writer.document.getCurrentController().getViewCursor()

      try:
          out_str = ResetCoding(cursor.getString())
          return out_str
      except:
          return ""
       
      

  ##
  # @brief カーソルの位置を取得する関数
  # @param self
  # @return カーソル位置のX座標、Y座標(単位はmm)
  #
  def oCurrentCursorPosition(self):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      oCurPos = cursor.getPosition()
      return oCurPos.X, oCurPos.Y

  ##
  # @brief カーソルをドキュメントの先頭に移動させる関数
  # @param self
  # @param sel Trueならば移動範囲を選択
  #
  def gotoStart(self, sel):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      cursor.gotoStart(sel)

  ##
  # @brief カーソルをドキュメントの最後尾に移動させる関数
  # @param self
  # @param sel Trueならば移動範囲を選択
  #
  def gotoEnd(self, sel):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      cursor.gotoEnd(sel)

  ##
  # @brief カーソルを行の先頭に移動させる関数
  # @param self
  # @param sel Trueならば移動範囲を選択
  #
  def gotoStartOfLine(self, sel):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      cursor.gotoStartOfLine(sel)

  ##
  # @brief カーソルの行の最後尾に移動させる関数
  # @param self
  # @param sel Trueならば移動範囲を選択
  #
  def gotoEndOfLine(self, sel):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      cursor.gotoEndOfLine(sel)

  
  
  
      

  ##
  # @brief 文字数移動する関数
  # @param self
  # @param diff 移動する文字数
  #
  def MoveCharacter(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      if diff > 0:
          cursor.goRight(diff,self.MovementType)
          if self.MovementType == False:
              cursor.collapseToEnd()
      else:
          cursor.goLeft(-diff,self.MovementType)
          if self.MovementType == False:
              cursor.collapseToStart()
          
  ##
  # @brief 単語数移動する関数
  # @param self
  # @param diff 移動する単語数
  #
  def MoveWord(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      for i in range(0, diff):
          if diff > 0:
              cursor.gotoNextWord(self.MovementType)
              if self.MovementType == False:
                  cursor.collapseToEnd()
          else:
              cursor.gotoPreviousWord(self.MovementType)
              if self.MovementType == False:
                  cursor.collapseToStart()

  ##
  # @brief 行数移動する関数
  # @param self
  # @param diff 移動する行数
  #
  def MoveLine(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      if diff > 0:
          cursor.goDown(diff,self.MovementType)
          if self.MovementType == False:
              cursor.collapseToEnd()
      else:
          cursor.goUp(-diff,self.MovementType)
          if self.MovementType == False:
              cursor.collapseToStart()

  ##
  # @brief 段落数移動する関数
  # @param self
  # @param diff 移動する段落数
  #
  def MoveParagraph(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      for i in range(0, diff):
          if diff > 0:
              cursor.gotoNextParagraph(self.MovementType)
              if self.MovementType == False:
                  cursor.collapseToEnd()
          else:
              cursor.gotoPreviousParagraph(self.MovementType)
              if self.MovementType == False:
                  cursor.collapseToStart()


  ##
  # @brief 活性化処理用コールバック関数
  # @param self
  # @param ec_id
  #
  
  def onActivated(self, ec_id):
    self.fontSize = float(self.conf_fontSize[0])
    self.fontName = self.conf_fontName[0]
    if int(self.conf_Bold[0]) == 0:
        self.Bold = False
    else:
        self.Bold = True
    if int(self.conf_Italic[0]) == 0:
        self.Italic = False
    else:
        self.Italic = True
    self.Char_Red = int(self.conf_Char_Red[0])
    self.Char_Green = int(self.conf_Char_Green[0])
    self.Char_Blue = int(self.conf_Char_Blue[0])

    if int(self.conf_Underline[0]) == 0:
        self.Underline = False
    else:
        self.Underline = True
    if int(self.conf_Shadow[0]) == 0:
        self.Shadow = False
    else:
        self.Shadow = True
    if int(self.conf_Strikeout[0]) == 0:
        self.Strikeout = False
    else:
        self.Strikeout = True
    if int(self.conf_Contoured[0]) == 0:
        self.Contoured = False
    else:
        self.Contoured = True
    if int(self.conf_Emphasis[0]) == 0:
        self.Emphasis = False
    else:
        self.Emphasis = True

    self.Back_Red = int(self.conf_Back_Red[0])
    self.Back_Green = int(self.conf_Back_Green[0])
    self.Back_Blue = int(self.conf_Back_Blue[0])

    #self.file = open('text3.txt', 'w')
    
    return RTC.RTC_OK

  def onDeactivated(self, ec_id):
    #self.file.close()
    return RTC.RTC_OK

  
  ##
  # @brief 周期処理用コールバック関数
  # @param self
  # @param ec_id
  #
  
  def onExecute(self, ec_id):
    

    if self._m_fontSizeIn.isNew():
        data = self._m_fontSizeIn.read()
        self.fontSize = data.data

    if self._m_MovementTypeIn.isNew():
        data = self._m_MovementTypeIn.read()
        self.MovementType = data.data

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

    if self._m_Char_colorIn.isNew():
        data = self._m_Char_colorIn.read()
        self.Char_Red = data.data.r*255
        self.Char_Green = data.data.g*255
        self.Char_Blue = data.data.b*255

    if self._m_ItalicIn.isNew():
        data = self._m_ItalicIn.read()
        self.Italic = data.data
    
    

    if self._m_BoldIn.isNew():
        data = self._m_BoldIn.read()
        self.Bold = data.data

    if self._m_UnderlineIn.isNew():
        data = self._m_UnderlineIn.read()
        self.Underline = data.data

    if self._m_ShadowIn.isNew():
        data = self._m_ShadowIn.read()
        self.Shadow = data.data

    if self._m_StrikeoutIn.isNew():
        data = self._m_StrikeoutIn.read()
        self.Strikeout = data.data

    if self._m_ContouredIn.isNew():
        data = self._m_ContouredIn.read()
        self.Contoured = data.data

    if self._m_EmphasisIn.isNew():
        data = self._m_EmphasisIn.read()
        self.Emphasis = data.data

    if self._m_Back_colorIn.isNew():
        data = self._m_Back_colorIn.read()
        self.Back_Red = data.data.r*255
        self.Back_Green = data.data.g*255
        self.Back_Blue = data.data.b*255

    

    if self._m_wordIn.isNew():
        data = self._m_wordIn.read()
        
        #t1_ = OpenRTM_aist.Time()
        self.SetWord(data.data)
        #t2_ = OpenRTM_aist.Time()
        #self.file.write(str((t2_-t1_).getTime().toDouble())+"\n")


    OpenRTM_aist.setTimestamp(self._d_m_selWord)
    self._d_m_selWord.data = str(self.GetWord())
    self._m_selWordOut.write()
        

    return RTC.RTC_OK

  
  ##
  # @brief 終了処理用コールバック関数
  # @param self
  # @param ec_id
  #
  
  def on_shutdown(self, ec_id):
      OOoRTC.writer_comp = None
      return RTC.RTC_OK



##
# @brief コンポーネントを活性化してWriterの操作を開始する関数
#

def Start():
    
    if OOoRTC.writer_comp:
        OOoRTC.writer_comp.m_activate()

##
# @brief コンポーネントを不活性化してWriterの操作を終了する関数
#

def Stop():
    
    if OOoRTC.writer_comp:
        OOoRTC.writer_comp.m_deactivate()


##
# @brief コンポーネントの実行周期を設定する関数
#

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
# @brief RTCをマネージャに登録する関数
#
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
# @brief RTC起動の関数
#

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

    
    MyMsgBox('',OOoRTC.SetCoding('RTCを起動しました','utf-8'))


    
    
    return None




##
# @brief メッセージボックス表示の関数
# @param title ウインドウのタイトル
# @param message 表示する文章
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
#

def MyMsgBox(title, message):
    try:
        m_bridge = Bridge()
    except:
        return
    m_bridge.run_infodialog(title, message)


##
# @brief OpenOfficeを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
#

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
# @brief OpenOffice Writerを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部を改変
#

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
