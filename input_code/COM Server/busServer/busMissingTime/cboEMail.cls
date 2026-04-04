VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cboEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Project:   TRIS Time to Pay
'Author:
'Date:
'Description:
Option Explicit

Private Const mstrCMODULENAME As String = "cboEMail"

Private m_sFrom As String
Private m_sTo As String
Private m_sCC As String
Private m_sBCC As String
Private m_sSubject As String
Private m_sTextBody As String
Private m_sHTMLBody As String

Public Property Let From(ByVal v_sFrom As String)
    m_sFrom = v_sFrom
End Property
Public Property Get From() As String
    From = m_sFrom
End Property

Public Property Let To_(ByVal v_sTo As String)
    m_sTo = v_sTo
End Property
Public Property Get To_() As String
    To_ = m_sTo
End Property

Public Property Let CC(ByVal v_sCC As String)
    m_sCC = v_sCC
End Property
Public Property Get CC() As String
    CC = m_sCC
End Property

Public Property Let BCC(ByVal v_sBCC As String)
    m_sBCC = v_sBCC
End Property
Public Property Get BCC() As String
    BCC = m_sBCC
End Property

Public Property Let Subject(ByVal v_sSubject As String)
    m_sSubject = v_sSubject
End Property
Public Property Get Subject() As String
    Subject = m_sSubject
End Property

Public Property Let TextBody(ByVal v_sTextBody As String)
    m_sTextBody = v_sTextBody
End Property
Public Property Get TextBody() As String
    TextBody = m_sTextBody
End Property

Public Property Let HTMLBody(ByVal v_sHTMLBody As String)
    m_sHTMLBody = v_sHTMLBody
End Property
Public Property Get HTMLBody() As String
    HTMLBody = m_sHTMLBody
End Property
