
{
  TdxSpreadSheet 扩展
  在dxSpreadSheet控件的滚动条左边分隔条处增加右键菜单，弹出显示当前的的所有工作表

  版权所有 (c) 研究  QQ:71051699

}
unit dxSpreadSheet.Helper;

interface

uses
  System.Classes,
  dxSpreadSheet,
  dxSpreadSheetCore,
  dxSpreadSheetPopupMenu;

type
  TdxSpreadSheetHelper = class helper for TdxSpreadSheet
  public
    procedure InitSplitterPopupMenu;
  end;

  TdxSpreadSheetPageControlViewInfoEx = class(TdxSpreadSheetPageControlViewInfo)
  protected
    function CreateSplitterCellViewInfo: TdxSpreadSheetPageControlSplitterCellViewInfo; override;
  end;

  TdxSpreadSheetPageControlHelper = class helper for TdxSpreadSheetPageControl
  public
    procedure CreateViewInfo(AViewInfo: TdxSpreadSheetPageControlViewInfoEx);
  end;

  TdxSpreadSheetBuiltInSheetsPopupMenu = class(TdxSpreadSheetCustomPopupMenu)
  private
    procedure MenuItemClick(Sender: TObject);
    function AddMenuItem(ACaption: String; ACommandID: Word; const AImageResName: string = ''; AEnabled: Boolean = True;
      AParent: TComponent = nil): TComponent; overload;
    procedure SetChecked(AItem: TComponent; AValue: Boolean); virtual;
  protected
    procedure PopulateMenuItems; override;
  public

  end;

  TdxSpreadSheetPageControlSplitterCellViewInfoEx = class(TdxSpreadSheetPageControlSplitterCellViewInfo)
  protected
    function GetPopupMenuClass(AHitTest: TdxSpreadSheetCustomHitTest): TComponentClass; override;
  public

  end;

implementation

{ TdxSpreadSheetPageControlViewInfoEx }

function TdxSpreadSheetPageControlViewInfoEx.CreateSplitterCellViewInfo
  : TdxSpreadSheetPageControlSplitterCellViewInfo;
begin
  Result := TdxSpreadSheetPageControlSplitterCellViewInfoEx.Create(Self);
end;

{ TdxSpreadSheetPageControlHelper }

procedure TdxSpreadSheetPageControlHelper.CreateViewInfo(AViewInfo: TdxSpreadSheetPageControlViewInfoEx);
begin
  with Self do
  begin
    FViewInfo.Free;
    FViewInfo := AViewInfo;
  end;
end;

{ TdxSpreadSheetBuiltInSheetsPopupMenu }

function TdxSpreadSheetBuiltInSheetsPopupMenu.AddMenuItem(ACaption: String; ACommandID: Word;
  const AImageResName: string; AEnabled: Boolean; AParent: TComponent): TComponent;
begin
  Result := Adapter.Add(ACaption, MenuItemClick, ACommandID, -1, AEnabled, 0, AParent);
end;

procedure TdxSpreadSheetBuiltInSheetsPopupMenu.MenuItemClick(Sender: TObject);
var
  AItem: TComponent;
begin
  AItem := TComponent(Sender);
  SetChecked(AItem, True);
  SpreadSheet.BeginUpdate();
  try
    SpreadSheet.ActiveSheet := SpreadSheet.Sheets[AItem.Tag];
  finally
    SpreadSheet.EndUpdate;
  end;
end;

procedure TdxSpreadSheetBuiltInSheetsPopupMenu.PopulateMenuItems;
var
  AItem: TComponent;
  I: Integer;
  ASheetCaption: String;
begin
  inherited;
  SpreadSheet.BeginUpdate();
  try
    for I := 0 to SpreadSheet.SheetCount - 1 do
    begin
      ASheetCaption := SpreadSheet.Sheets[I].Caption;
      if SpreadSheet.Sheets[I].Visible then
      begin
        AItem := AddMenuItem(ASheetCaption, I);
        SetChecked(AItem, SpreadSheet.ActiveSheetIndex = I);
      end;
    end;
  finally
    SpreadSheet.EndUpdate;
  end;
end;

procedure TdxSpreadSheetBuiltInSheetsPopupMenu.SetChecked(AItem: TComponent; AValue: Boolean);
begin
  inherited;
  Adapter.SetChecked(AItem, AValue);
end;

{ TdxSpreadSheetPageControlSplitterCellViewInfoEx }

function TdxSpreadSheetPageControlSplitterCellViewInfoEx.GetPopupMenuClass(AHitTest: TdxSpreadSheetCustomHitTest)
  : TComponentClass;
begin
  Result := TdxSpreadSheetBuiltInSheetsPopupMenu;
end;

{ TdxSpreadSheetHelper }

procedure TdxSpreadSheetHelper.InitSplitterPopupMenu;
var
  AViewInfo: TdxSpreadSheetPageControlViewInfoEx;
begin
  AViewInfo := TdxSpreadSheetPageControlViewInfoEx.Create(Self.PageControl);
  Self.PageControl.CreateViewInfo(AViewInfo);
end;

end.