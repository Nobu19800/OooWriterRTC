#!/bin/env python
# -*- encoding: utf-8 -*-

##
#
# @file WriterControl.py

import optparse
import sys,os,platform
import re
import time
import random
import commands
import math




from os.path import expanduser
sv = sys.version_info


if os.name == 'posix':
    home = expanduser("~")
    sys.path += [home+'/OOoRTC', home+'/OOoRTC/WriterIDL', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages/rtctree/rtmidl']
elif os.name == 'nt':
    sys.path += ['.\\OOoRTC', '.\\OOoRTC\\WriterIDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages', 'C:\\Python2' + str(sv[1]) + '\\Lib\\site-packages\\OpenRTM_aist\\RTM_IDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages\\rtctree\\rtmidl']




import RTC
import OpenRTM_aist

from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming

import Writer_idl
from omniORB import PortableServer
import Writer, Writer__POA



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
# @brief 入力文字列をWriterのドキュメント上で文字化けしない文字コードで文字列を返す
# @param m_str 変換前の文字列
# @param m_code 変換前の文字コード
# @return 変換後の文字列
#
def SetCoding(m_str, m_code):
    if os.name == 'posix':
        if m_code == "utf-8":
            return m_str
        else:
            try:
                return m_str.decode(m_code).encode("utf-8")
            except:
                return ""
    elif os.name == 'nt':
        try:
            return m_str.decode(m_code).encode('cp932')
        except:
            return "" 



##
# @class WordControl
# @brief 文書作成ソフトを操作するためのRTCのクラス
#

class WriterControl(OpenRTM_aist.DataFlowComponentBase):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param manager マネージャーオブジェクト
    #
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
    

    

    self.fontSize = 16
    self.fontName = "ＭＳ 明朝"
    self.bold = False
    self.italic = False
    self.char_Red = 0
    self.char_Green = 0
    self.char_Blue = 0
    self.movementType = False

    self.underline = False
    self.shadow = False
    self.strikeout = False
    self.contoured = False
    self.emphasis = False

    self.back_Red = 255
    self.back_Green = 255
    self.back_Blue = 255


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
  
  def mSetRate(self, rate):
      m_ec = self.get_owned_contexts()
      m_ec[0].set_rate(rate)

  ##
  # @brief 活性化するための関数
  # @param self 
  #
  def mActivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].activate_component(self._objref)

  ##
  # @brief 不活性化するための関数
  # @param self 
  #
  def mDeactivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].deactivate_component(self._objref)
      

  ##
  # @brief 初期化処理用コールバック関数
  # @param self 
  # @return RTC::ReturnCode_t
  def onInitialize(self):
    

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
    self.bindParameter("bold", self.conf_Bold, "0")
    self.bindParameter("italic", self.conf_Italic, "0")
    self.bindParameter("char_Red", self.conf_Char_Red, "0")
    self.bindParameter("char_Blue", self.conf_Char_Blue, "0")
    self.bindParameter("char_Green", self.conf_Char_Green, "0")
    self.bindParameter("code", self.conf_Code, "utf-8")

    self.bindParameter("underline", self.conf_Underline, "0")
    self.bindParameter("shadow", self.conf_Shadow, "0")
    self.bindParameter("strikeout", self.conf_Strikeout, "0")
    self.bindParameter("contoured", self.conf_Contoured, "0")
    self.bindParameter("emphasis", self.conf_Emphasis, "0")

    self.bindParameter("back_Red", self.conf_Back_Red, "255")
    self.bindParameter("back_Blue", self.conf_Back_Blue, "255")
    self.bindParameter("back_Green", self.conf_Back_Green, "255")

    
        
    
    return RTC.RTC_OK

  ##
  # @brief 文字書き込みの関数
  # @param self 
  # @param m_str 書き込む文字列
  #

  def setWord(self, m_str):
      pass


  ##
  # @brief カーソル位置の文字取得の関数
  # @param self
  # @return カーソル位置の文字列
  #

  def getWord(self):
      return ""

  ##
  # @brief 活性化処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onActivated(self, ec_id):
    
    
    self.fontSize = float(self.conf_fontSize[0])
    self.fontName = self.conf_fontName[0]
    if int(self.conf_Bold[0]) == 0:
        self.bold = False
    else:
        self.bold = True
    if int(self.conf_Italic[0]) == 0:
        self.italic = False
    else:
        self.italic = True
    self.char_Red = int(self.conf_Char_Red[0])
    self.char_Green = int(self.conf_Char_Green[0])
    self.char_Blue = int(self.conf_Char_Blue[0])

    if int(self.conf_Underline[0]) == 0:
        self.underline = False
    else:
        self.underline = True
    if int(self.conf_Shadow[0]) == 0:
        self.shadow = False
    else:
        self.shadow = True
    if int(self.conf_Strikeout[0]) == 0:
        self.strikeout = False
    else:
        self.strikeout = True
    if int(self.conf_Contoured[0]) == 0:
        self.contoured = False
    else:
        self.contoured = True
    if int(self.conf_Emphasis[0]) == 0:
        self.Emphasis = False
    else:
        self.Emphasis = True

    self.back_Red = int(self.conf_Back_Red[0])
    self.back_Green = int(self.conf_Back_Green[0])
    self.back_Blue = int(self.conf_Back_Blue[0])

    
    
    return RTC.RTC_OK


  ##
  # @brief 不活性化処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  def onDeactivated(self, ec_id):
    
    return RTC.RTC_OK


  ##
  # @brief 周期処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onExecute(self, ec_id):
    
    
    

    if self._m_fontSizeIn.isNew():
        data = self._m_fontSizeIn.read()
        self.fontSize = data.data

    if self._m_MovementTypeIn.isNew():
        data = self._m_MovementTypeIn.read()
        self.movementType = data.data

    if self._m_wsCharacterIn.isNew():
        data = self._m_wsCharacterIn.read()
        self.moveCharacter(data.data)

    if self._m_wsWordIn.isNew():
        data = self._m_wsWordIn.read()
        self.moveWord(data.data)

    if self._m_wsLineIn.isNew():
        data = self._m_wsLineIn.read()
        self.moveLine(data.data)

    if self._m_wsParagraphIn.isNew():
        data = self._m_wsParagraphIn.read()
        self.moveParagraph(data.data)

    if self._m_wsWindowIn.isNew():
        data = self._m_wsWindowIn.read()
        self.moveWindow(data.data)


    if self._m_wsScreenIn.isNew():
        data = self._m_wsScreenIn.read()
        self.moveScreen(data.data)

    if self._m_Char_colorIn.isNew():
        data = self._m_Char_colorIn.read()
        self.char_Red = data.data.r*255
        self.char_Green = data.data.g*255
        self.char_Blue = data.data.b*255

    if self._m_ItalicIn.isNew():
        data = self._m_ItalicIn.read()
        self.italic = data.data
    
    

    if self._m_BoldIn.isNew():
        data = self._m_BoldIn.read()
        self.bold = data.data

    if self._m_UnderlineIn.isNew():
        data = self._m_UnderlineIn.read()
        self.underline = data.data

    if self._m_ShadowIn.isNew():
        data = self._m_ShadowIn.read()
        self.shadow = data.data

    if self._m_StrikeoutIn.isNew():
        data = self._m_StrikeoutIn.read()
        self.strikeout = data.data

    if self._m_ContouredIn.isNew():
        data = self._m_ContouredIn.read()
        self.contoured = data.data

    if self._m_EmphasisIn.isNew():
        data = self._m_EmphasisIn.read()
        self.emphasis = data.data

    if self._m_Back_colorIn.isNew():
        data = self._m_Back_colorIn.read()
        self.back_Red = data.data.r*255
        self.back_Green = data.data.g*255
        self.back_Blue = data.data.b*255

    

    if self._m_wordIn.isNew():
        data = self._m_wordIn.read()
        
        #t1_ = OpenRTM_aist.Time()
        self.setWord(data.data)
        #t2_ = OpenRTM_aist.Time()
        #self.file.write(str((t2_-t1_).getTime().toDouble())+"\n")


    OpenRTM_aist.setTimestamp(self._d_m_selWord)
    self._d_m_selWord.data = str(self.getWord())
    self._m_selWordOut.write()
        

    return RTC.RTC_OK

  
  ##
  # @brief 終了処理用コールバック関数
  # @param self 
  # @param 
  # @return RTC::ReturnCode_t
  
  def onFinalize(self):
      return RTC.RTC_OK


  ##
  # @brief カーソルの位置を取得する関数
  # @param self 
  # @return カーソル位置のX座標、Y座標(単位はmm)
  #
  def oCurrentCursorPosition(self):
      return 0,0

  ##
  # @brief カーソルをドキュメントの先頭に移動させる関数
  # @param self 
  # @param sel Trueならば移動範囲を選択
  #
  def gotoStart(self, sel):
      pass

  ##
  # @brief カーソルをドキュメントの最後尾に移動させる関数
  # @param self 
  # @param sel Trueならば移動範囲を選択
  #
  def gotoEnd(self, sel):
      pass

  ##
  # @brief カーソルを行の先頭に移動させる関数
  # @param self 
  # @param sel Trueならば移動範囲を選択
  #
  def gotoStartOfLine(self, sel):
      pass

  ##
  # @brief カーソルの行の最後尾に移動させる関数
  # @param self 
  # @param sel Trueならば移動範囲を選択
  #
  def gotoEndOfLine(self, sel):
      pass

  
  
  
      

  ##
  # @brief 文字数移動する関数
  # @param self 
  # @param diff 移動する文字数
  #
  def moveCharacter(self, diff):
      pass
          
  ##
  # @brief 単語数移動する関数
  # @param self 
  # @param diff 移動する単語数
  #
  def moveWord(self, diff):
      pass

  ##
  # @brief 行数移動する関数
  # @param self 
  # @param diff 移動する行数
  #
  def moveLine(self, diff):
      pass

  ##
  # @brief 段落数移動する関数
  # @param self 
  # @param diff 移動する段落数
  #
  def moveParagraph(self, diff):
      pass

  ##
  # @brief ウインドウ数移動する関数
  # @param self 
  # @param diff 移動する段落数
  #
  def moveWindow(self, diff):
      pass

  ##
  # @brief スクリーン数移動する関数
  # @param self 
  # @param diff 移動する段落数
  #
  def moveScreen(self, diff):
      pass
