object Frame1: TFrame1
  Left = 0
  Top = 0
  Width = 1488
  Height = 287
  Color = 8632963
  ParentColor = False
  TabOrder = 0
  object Label1: TLabel
    Left = 319
    Top = 7
    Width = 85
    Height = 13
    Caption = 'Owner Search:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label2: TLabel
    Left = 23
    Top = 7
    Width = 75
    Height = 13
    Caption = 'Acct Search:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label3: TLabel
    Left = 610
    Top = 7
    Width = 94
    Height = 13
    Caption = 'Address Search:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object Label4: TLabel
    Left = 946
    Top = 7
    Width = 83
    Height = 13
    Caption = 'Street Search:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object sbAcctSort: TSpeedButton
    Left = 103
    Top = 23
    Width = 57
    Height = 22
    Caption = 'Acct:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333303333
      333333333337F33333333333333033333333333333373F333333333333090333
      33333333337F7F33333333333309033333333333337373F33333333330999033
      3333333337F337F33333333330999033333333333733373F3333333309999903
      333333337F33337F33333333099999033333333373333373F333333099999990
      33333337FFFF3FF7F33333300009000033333337777F77773333333333090333
      33333333337F7F33333333333309033333333333337F7F333333333333090333
      33333333337F7F33333333333309033333333333337F7F333333333333090333
      33333333337F7F33333333333300033333333333337773333333}
    NumGlyphs = 2
    ParentFont = False
    OnClick = sbAcctSortClick
  end
  object sbOwnerSort: TSpeedButton
    Left = 399
    Top = 23
    Width = 65
    Height = 22
    Caption = 'Owner:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333303333
      333333333337F33333333333333033333333333333373F333333333333090333
      33333333337F7F33333333333309033333333333337373F33333333330999033
      3333333337F337F33333333330999033333333333733373F3333333309999903
      333333337F33337F33333333099999033333333373333373F333333099999990
      33333337FFFF3FF7F33333300009000033333337777F77773333333333090333
      33333333337F7F33333333333309033333333333337F7F333333333333090333
      33333333337F7F33333333333309033333333333337F7F333333333333090333
      33333333337F7F33333333333300033333333333337773333333}
    NumGlyphs = 2
    ParentFont = False
    OnClick = sbOwnerSortClick
  end
  object sbStrNumSort: TSpeedButton
    Left = 839
    Top = 23
    Width = 55
    Height = 22
    Caption = 'Num:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333303333
      333333333337F33333333333333033333333333333373F333333333333090333
      33333333337F7F33333333333309033333333333337373F33333333330999033
      3333333337F337F33333333330999033333333333733373F3333333309999903
      333333337F33337F33333333099999033333333373333373F333333099999990
      33333337FFFF3FF7F33333300009000033333337777F77773333333333090333
      33333333337F7F33333333333309033333333333337F7F333333333333090333
      33333333337F7F33333333333309033333333333337F7F333333333333090333
      33333333337F7F33333333333300033333333333337773333333}
    NumGlyphs = 2
    ParentFont = False
    OnClick = sbStrNumSortClick
  end
  object sbStrNameSort: TSpeedButton
    Left = 1031
    Top = 23
    Width = 65
    Height = 22
    Caption = 'Street:'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    Glyph.Data = {
      76010000424D7601000000000000760000002800000020000000100000000100
      04000000000000010000120B0000120B00001000000000000000000000000000
      800000800000008080008000000080008000808000007F7F7F00BFBFBF000000
      FF0000FF000000FFFF00FF000000FF00FF00FFFF0000FFFFFF00333333303333
      333333333337F33333333333333033333333333333373F333333333333090333
      33333333337F7F33333333333309033333333333337373F33333333330999033
      3333333337F337F33333333330999033333333333733373F3333333309999903
      333333337F33337F33333333099999033333333373333373F333333099999990
      33333337FFFF3FF7F33333300009000033333337777F77773333333333090333
      33333333337F7F33333333333309033333333333337F7F333333333333090333
      33333333337F7F33333333333309033333333333337F7F333333333333090333
      33333333337F7F33333333333300033333333333337773333333}
    NumGlyphs = 2
    ParentFont = False
    OnClick = sbStrNameSortClick
  end
  object frm1DbGrid: TDBGrid
    Left = 8
    Top = 55
    Width = 1472
    Height = 210
    DataSource = dsCurrentOwners
    FixedColor = clMoneyGreen
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    ParentFont = False
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = [fsBold]
    OnDblClick = frm1DbGridDblClick
    Columns = <
      item
        Expanded = False
        FieldName = 'ownerID'
        Visible = False
      end
      item
        Expanded = False
        FieldName = 'houseAcct'
        Title.Caption = 'Acct'
        Width = 45
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Owner'
        Width = 235
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'legal'
        Title.Caption = 'Legal'
        Width = 60
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Phone'
        Width = 80
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'AltPhone'
        Title.Caption = 'Alt Phone'
        Width = 80
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Offsite'
        Width = 60
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Address'
        Width = 100
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Zip'
        Width = 65
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Lot'
        Width = 20
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Block'
        Width = 20
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Section'
        Width = 30
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'StreetNumber'
        Title.Caption = 'Street #'
        Width = 45
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'StreetName'
        Title.Caption = 'Str Name'
        Width = 80
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'driveRoute'
        Title.Caption = 'Drive Rte'
        Width = 24
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'greeting'
        Title.Caption = 'Greeting'
        Width = 175
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'mailName'
        Title.Caption = 'Mail Name'
        Width = 220
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'closeDate'
        Title.Caption = 'Close Date'
        Width = 61
        Visible = True
      end>
  end
  object eOwnersOwner: TEdit
    Left = 319
    Top = 23
    Width = 75
    Height = 21
    TabOrder = 1
    OnChange = eOwnersOwnerChange
    OnEnter = EditColorChangeOnEnter
    OnExit = EditColorChangeOnExit
  end
  object sbCurrentOwnersFrame: TStatusBar
    Left = 0
    Top = 268
    Width = 1488
    Height = 19
    Panels = <
      item
        Width = 150
      end
      item
        Width = 150
      end>
  end
  object Edit1: TEdit
    Left = 32
    Top = 24
    Width = 65
    Height = 21
    TabOrder = 3
    OnChange = eOwnersHouseAcctChange
    OnEnter = EditColorChangeOnEnter
    OnExit = EditColorChangeOnExit
  end
  object Edit2: TEdit
    Left = 944
    Top = 24
    Width = 75
    Height = 21
    TabOrder = 4
    OnChange = eOwnersStreetSearchChange
    OnEnter = EditColorChangeOnEnter
    OnExit = EditColorChangeOnExit
  end
  object Edit3: TEdit
    Left = 608
    Top = 24
    Width = 75
    Height = 21
    TabOrder = 5
    OnChange = eOwnersAddressSearchChange
    OnEnter = EditColorChangeOnEnter
    OnExit = EditColorChangeOnExit
  end
  object frame1AdoTableCurrentOwners: TADOTable
    Connection = frame1AdoConnection
    CursorType = ctStatic
    Filtered = True
    TableName = 'CurrentOwners'
    Left = 503
    Top = 143
  end
  object frame1AdoConnection: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Mandy.SSCA' +
      '-FRONTOFFIC\Documents\SSCA_ACDR_Rev171011.mdb;Persist Security I' +
      'nfo=False;'
    LoginPrompt = False
    Mode = cmShareDenyNone
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 543
    Top = 143
  end
  object dsCurrentOwners: TDataSource
    DataSet = frame1AdoTableCurrentOwners
    Left = 503
    Top = 183
  end
end
