/* USER1_B_PROJECT_0.J
 */


; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)
; b_project_0

ErrMsg := "";
oleserver := "excel.application";
goto begin;

; General exception raiser
@generalerror:
  if !excel.visible then excel.visible := true;
  MsgErr(FormatStr(GetLanText(-9688), ErrMsg));
  Halt;

; Error in oleserver
@oleservererror:
  if !excel.visible then excel.visible := true;
  MsgErr(FormatStr(GetLanText(-9687), ErrMsg));
  Halt;

; Start processing
@begin:
excel := start(oleserver);
if !IsIDispatch(excel) then
{
  ErrMsg := oleserver;
  goto oleservererror;
}
else
{
  oleversion := StrToNum(GetParam("OFFICE"), 0);
  oleversion := StrToNum(excel.Version, oleversion);
  excel.visible := True;
};

; Open a temporary template file for calculations
TemplateFile := FileSearch("%ESTIMATION_TEMPLATE%.XLT", "%PATH_DATA%");
if templatefile = "" || !FileExists(templatefile) then
{
  ErrMsg := "Cannot find template <%ESTIMATION_TEMPLATE%.XLT> in <%PATH_DATA%>!";
  goto generalerror;
}
Template := excel.workbooks.add(TemplateFile);
if !IsIDispatch(Template) then
{
  ErrMsg := "Open <" + TemplateFile + "> failed!";
  goto generalerror;
}
Template.Author := "%DB_USERDESC%";
CostSheet := Template.WorkSheets["Cost"];
DataSheet := Template.WorkSheets["Data"];
HelpSheet := Template.WorkSheets["Help"];

; Translated formulas
TemplateFormulaCell := DataSheet.Range["Formula"];
TemplateFormula := TemplateFormulaCell.FormulaR1C1;
BracketPos := StrPos("(", TemplateFormula);
SumFormulaText := SubStr(TemplateFormula, 2, BracketPos - 2);
RId := SubStr(TemplateFormula, BracketPos + 1, 1);
CId := SubStr(TemplateFormula, BracketPos + 2, 1);
LBr := SubStr(TemplateFormula, BracketPos + 3, 1);
RBr := SubStr(TemplateFormula, BracketPos + 5, 1);

; Cell formats
CellSizeFormat    := "#" + "%THOUSANDSEP%" + "##0";
CellCostFormat    := "#" + "%THOUSANDSEP%" + "##0" + "%DECIMALSEP%" + "00";
CellAreaFormat    := "#" + "%THOUSANDSEP%" + "##0" + "%DECIMALSEP%" + "00";
CellPriceFormat   := "#" + "%THOUSANDSEP%" + "##0"; 
CellCountFormat   := "#" + "%THOUSANDSEP%" +"##0[$x]";
CellFactorFormat  := "#" + "%THOUSANDSEP%" + "##0" + "%DECIMALSEP%" + "00"; 
CellPercentFormat := "##0" + "%DECIMALSEP%" + "00%"; 
CellLineFormat    := "0";
CellTextFormat    := "@";

; Default template row & column indexes
Count := 0;
Range := 0;
ColCT := CostSheet.Range["CostTime"].Column;
ColC1 := CostSheet.Range["CostLoss"].Column;
ColC2 := CostSheet.Range["CostDiscount"].Column;
ColC7 := CostSheet.Range["CostSystem"].Column;
ColC3 := CostSheet.Range["CostFactor"].Column;
ColC6 := CostSheet.Range["CostCharge"].Column;
ColC4 := CostSheet.Range["CostProfit"].Column;
ColC5 := CostSheet.Range["CostRebate"].Column;
ColId := CostSheet.Range["PriceBlocks"].Column;
RowId := CostSheet.Range["PriceBlocks"].Row;
Color := DataSheet.Range["HeadFormat"].Interior.Color;
LossF := DataSheet.Range["Fixed"].Value;
LossA := DataSheet.Range["Absolute"].Value;
LossR := DataSheet.Range["Real"].Value;
bList := Strings.Create();                                  /*artikel+verliestype*/
cList := Strings.Create();                                  /*rowid*/
pList := Strings.Create();                                  /*db_res_cost【price】*/
sList := Strings.Create();                                  /*rowid*/
/*用于记录E_PROJECT.J计算所用数据，当前记录顺序：窗号-类型-樘数*/
/*hList:=Strings.Create();*/
help_cnt:=4;


RowId := RowId + 1;                                         /*更改标题单元格线型格式*/
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColCT];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC1];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC2];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC7];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC3];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC6];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC4];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC5];
CurrentCell.Borders.linestyle:=0;

RowId := RowId + 1;
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColCT];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC1];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC2];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC7];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC3];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC6];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC4];
CurrentCell.Borders.linestyle:=0;
CurrentCell := CostSheet.Cells[RowId][ColC5];
CurrentCell.Borders.linestyle:=0;


RowId_G:=0;                                                 /*玻璃行号*/
glass_price:=0;                                             /*玻璃价格标志位，*/
                                                            /*1-表示已记录玻璃行号*/
                                                            
%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail
; 

a_link:="";
z_pg:=pricegroups.create();
z_pg.code.group:="A";
z_pg.code.block:=@%DB_COST_ARTICLE%;
if z_pg.find() then
{
	a_link:=z_pg.link;
}
else
{
	msgbox("no article block "+inttostr(z_pg.code.block)+" found!");
}


RowId := RowId + 1;
Range := Range + 1;

TempValue    := %IF{%EVAL{@%DB_RES_COST%>0},"@%DB_RES_COST%","0"};

if a_link<>"" then                                          /*记录价格块或跟随块*/
	bList.Add(a_link+"@%DB_COST_LOSSTYPE%");
else
	bList.Add("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%");
cList.Add(IntToStr(RowId));                                 /*记录行号*/
pList.Add(StrReplace("TempValue",".","%DECIMALSEP%"));      /*记录价格*/
sList.Add(inttostr(RowId));                                 /*记录行号*/

/* 记录玻璃行号并设置玻璃价格标志位1*/
if @%DB_COST_ARTICLE%=20 then
{
    RowId_G:=RowId;
    glass_price:=1;
}

CostSheet.Cells[RowId][1].value:=rowid-3;                   /*输出序号*/
CostSheet.Cells[RowId][1].Borders.linestyle:=0;

; Priceblock description
TempValue   := "@%DB_RES_DESC%";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellTextFormat;
CurrentCell.Font.Bold := False;
CurrentCell.Borders.linestyle:=0;

; Priceblock cost
TempValue   := %IF{%EVAL{%ALLOW_COST_RATES%=1},%IF{%EVAL{@%DB_RES_COST%>0},@%DB_RES_COST%,0},0};
CurrentCell := CostSheet.Cells[RowId][ColCT];
CurrentCell.Value := TempValue;
CurrentCell.HorizontalAlignment := 1;
CurrentCell.NumberFormat := CellCostFormat;
CurrentCell.Font.Italic := %IF{%EVAL{@%DB_RES_COST%>0},False,True};
CurrentCell.Borders.linestyle:=0;

; Priceblock loss
TempValue   := %IF{%EVAL{%ALLOW_COST_RATES%=1},@%DB_RES_LOSS%/100,0};
CurrentCell := CostSheet.Cells[RowId][ColC1];
CurrentCell.Value := TempValue;
CurrentCell.HorizontalAlignment := 1;
CurrentCell.NumberFormat := CellPercentFormat;
/*CurrentCell.Interior.Color := Color;*/
CurrentCell.Borders.linestyle:=0;

; Priceblock discount
TempValue   := %IF{%EVAL{%ALLOW_COST_RATES%=1},@%DB_RES_DISCOUNT%/100,0};
CurrentCell := CostSheet.Cells[RowId][ColC2];
CurrentCell.Value := TempValue;
CurrentCell.HorizontalAlignment := 1;
CurrentCell.NumberFormat := CellPercentFormat;
/*CurrentCell.Interior.Color := Color;*/
CurrentCell.Borders.linestyle:=0;

; Priceblock system
TempValue   := %IF{%EVAL{%ALLOW_COST_RATES%=1},@%DB_COST_RATION%,1};
CurrentCell := CostSheet.Cells[RowId][ColC7];
CurrentCell.Value := TempValue;
CurrentCell.HorizontalAlignment := 1;
CurrentCell.NumberFormat := CellFactorFormat;
/*CurrentCell.Interior.Color := Color;*/
CurrentCell.Borders.linestyle:=0;

; Priceblock factor
TempValue   := %IF{%EVAL{%ALLOW_COST_RATES%=1},@%DB_COST_FACTOR%,1};
CurrentCell := CostSheet.Cells[RowId][ColC3];
CurrentCell.Value := TempValue;
CurrentCell.HorizontalAlignment := 1;
CurrentCell.NumberFormat := CellFactorFormat;
/*CurrentCell.Interior.Color := Color;*/
CurrentCell.Borders.linestyle:=0;

; Priceblock charge
TempValue   := %IF{%EVAL{%ALLOW_COST_RATES%=1},@%DB_RES_CHARGE%/100,0};
CurrentCell := CostSheet.Cells[RowId][ColC6];
CurrentCell.Value := TempValue;
CurrentCell.HorizontalAlignment := 1;
CurrentCell.NumberFormat := CellPercentFormat;
/*CurrentCell.Interior.Color := Color;*/
CurrentCell.Borders.linestyle:=0;

; Priceblock profit
TempValue   := %IF{%EVAL{%ALLOW_COST_RATES%=1},@%DB_RES_PROFIT%/100,0};
CurrentCell := CostSheet.Cells[RowId][ColC4];
CurrentCell.Value := TempValue;
CurrentCell.HorizontalAlignment := 1;
CurrentCell.NumberFormat := CellPercentFormat;
/*CurrentCell.Interior.Color := Color;*/
CurrentCell.Borders.linestyle:=0;

; Priceblock rebate
TempValue   := %IF{%EVAL{%ALLOW_COST_RATES%=1},@%DB_RES_REFUND%/100,0};
CurrentCell := CostSheet.Cells[RowId][ColC5];
CurrentCell.Value := TempValue;
CurrentCell.HorizontalAlignment := 1;
CurrentCell.NumberFormat := CellPercentFormat;
/*CurrentCell.Interior.Color := Color;*/
CurrentCell.Borders.linestyle:=0;

%% break header
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Break header
; 

%% break footer
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Break footer
; 

%% detail footer
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail footer
; 

; Column unit price item
RowId := RowId + 1;
bList.Add("-1");
cList.Add(IntToStr(RowId));
sList.Add(inttostr(RowId));

; Column number of items
RowId := RowId + 1;
bList.Add("-2");
cList.Add(IntToStr(RowId));
sList.Add(inttostr(RowId));

; Column total price item
RowId := RowId + 1;
bList.Add("-3");
cList.Add(IntToStr(RowId));
sList.Add(inttostr(RowId));



recent_count:=count;                                        /*记录当前状态*/
recent_colid:=colid;
recent_cost_sheet:=CostSheet;


ColId := ColId + 12;                                        /*将ColId=2增加12，定位到单价列*/
Color := DataSheet.Range["CellFormat"].Interior.Color;

; Initialize prices project level
i := 0;
while (i < cList.Count-3) do                                /*cList最后三行为总价统计，忽略*/
{
  RowId       := StrToNum(cList.Strings[i]);                /*从cList中获取行号*/
  TempValue   := 0.0;                                       /*临时变量*/
  CurrentCell := CostSheet.Cells[RowId][ColId];             /*定位到当前单元格，单价列*/
  CurrentCell0 := CostSheet.Cells[RowId][ColId-1];          /*定位到相邻左侧单元格，单樘用量列*/
  CurrentCell.Value := TempValue;
  CurrentCell0.Value := TempValue;
  CurrentCell.Borders.linestyle := 0;
  i := i + 1;
};


recent_count:=count;                                        /*记录当前状态*/
recent_colid:=colid;
recent_cost_sheet:=CostSheet;

total_area:=0;                                              /*工程总面积，用于计算工程级价格块*/


