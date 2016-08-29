object FrmEligeFecha: TFrmEligeFecha
  Left = 0
  Top = 0
  BorderIcons = []
  BorderStyle = bsToolWindow
  Caption = 'Elegir Fecha'
  ClientHeight = 169
  ClientWidth = 147
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object DNFecha: TcxDateNavigator
    Left = 0
    Top = 0
    Width = 147
    Height = 129
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    TabOrder = 0
  end
  object btnAceptar: TNxButton
    Left = 8
    Top = 147
    Width = 57
    Height = 17
    Caption = '&Aceptar'
    ModalResult = 1
    TabOrder = 1
  end
  object btnCancelar: TNxButton
    Left = 82
    Top = 147
    Width = 57
    Height = 17
    Caption = '&Cancelar'
    ModalResult = 2
    TabOrder = 2
  end
end
