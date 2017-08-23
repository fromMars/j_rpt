Count:=recent_count;
ColId:=recent_colid;
CostSheet:=recent_cost_sheet;

CostSheet.Copy(CostSheet);
cnt := Template.WorkSheets.Count-3;

template_cost:=Template.WorkSheets[cnt];
template_cost.Name := "Cost_"+trim("%ASSEMBLY_TEXT%");
template_cost.Activate();

CostSheet:=template_cost;

CostSheet.Range["RateRows"].Delete;
ColId:=ColId-8;


recent_rowid:=-1;


curr_assembly:=getcurrentproject().projectdata.children[cnt-1];
;assembly name
curr_name:=curr_assembly.code;
costsheet.range["chuanghao"].value:="窗号："+curr_name;

;surface
curr_width:=curr_assembly.width;
curr_height:=curr_assembly.height;
curr_surface:=curr_width*curr_height/1000000;
costsheet.range["mianji"].value:=curr_surface;
a_fee_row:=0;

; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)  b_assembly_0.j
; 

colid:=6;
Count := Count + 1;
ColId := ColId + 1;
Color := DataSheet.Range["CellFormat"].Interior.Color;

; Initialize prices assembly level
i := 0;
while (i < cList.Count-3) do
{
  RowId       := StrToNum(cList.Strings[i]);
  TempValue   := 0.0;
  CurrentCell := CostSheet.Cells[RowId][ColId];
  CurrentCell.Value := TempValue;
  CurrentCell1 := CostSheet.Cells[RowId][ColId-1];
  CurrentCell1.Value := TempValue;
  CurrentCell0 := CostSheet.Cells[RowId][ColId-2];
  CurrentCell0.Value := TempValue;
  /*CurrentCell.NumberFormat := CellPriceFormat;*/
  CurrentCell.Font.Italic := True;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
  i := i + 1;
};

%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail
; 

/*
; Item counter
RowId       := 1;
TempValue   := Count;
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellLineFormat;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;*/

/*
; Item description
RowId       := 2;
TempValue   := "@%DB_COST_ID%";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellTextFormat;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;

; Item width
RowId       := 3;
TempValue   := "";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellTextFormat;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;

; Item height
RowId       := 4;
TempValue   := "";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellTextFormat;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;

; Item surface
RowId       := 5;
TempValue   := "";
CurrentCell := CostSheet.Cells[RowId][ColId];
CurrentCell.Value := TempValue;
CurrentCell.NumberFormat := CellTextFormat;
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;
*/


; Item price
/*RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%"+"@%DB_COST_RATION%"+"@%DB_COST_FACTOR%"+"@%DB_COST_RATIO%")]);*/
if @%DB_PIECE_ARTICLE%=19 || @%DB_PIECE_ARTICLE%=18 || @%DB_PIECE_ARTICLE%=8 ||@%DB_PIECE_ARTICLE%=9 || @%DB_PIECE_ARTICLE%=10 then
{
	RowId  := StrToNum(cList.Strings[bList.IndexOf("19"+"@%DB_PIECE_LOSSTYPE%")]);
}
else
{
	RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_PIECE_ARTICLE%"+"@%DB_PIECE_LOSSTYPE%")]);
}
if recent_rowid=-1 then
	recent_rowid:=rowid;

CellCT := "Indirect(address("+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColCT)+"))/"+pList.Strings[cList.IndexOf(IntToStr(RowId))];
CellC1 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC1)+"))";
CellC2 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC2)+"))";
CellC7 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC7)+"))";
CellC3 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC3)+"))";
CellC4 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC4)+"))";
CellC5 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC5)+"))";
CellC6 := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColC6)+"))";
if (StrToNum(StrReplace(pList.Strings[cList.IndexOf(IntToStr(RowId))],"%DECIMALSEP%","."),0) > 0) then
{
  CurrentCell := CostSheet.Cells[RowId][ColId];
  TempValue   := StrReplace("@%DB_PIECE_PRICE%/%ASSEMBLYCOUNT%",".","%DECIMALSEP%");
  if @%DB_PIECE_ARTICLE%=18 || @%DB_PIECE_ARTICLE%=8 ||@%DB_PIECE_ARTICLE%=9 || @%DB_PIECE_ARTICLE%=10 then
  {
	curr_profile_value:=currentcell.value;
	TempFormula := "=((((((((("+TempValue+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))+"+numtostr(curr_profile_value);
  }
  else
  {
	TempFormula := "=((((((((("+TempValue+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  }
  
  
  CurrentCell.Formula := TempFormula;
/*  CurrentCell.NumberFormat := CellPriceFormat;*/
  CurrentCell.Font.Italic := False;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
}
else
{
  TempValue   := StrReplace("@%DB_PIECE_PRICE%/%ASSEMBLYCOUNT%",".","%DECIMALSEP%");
  CurrentCell := CostSheet.Cells[RowId][ColId];

  if @%DB_PIECE_ARTICLE%=19 || @%DB_PIECE_ARTICLE%=18 || @%DB_PIECE_ARTICLE%=8 ||@%DB_PIECE_ARTICLE%=9 || @%DB_PIECE_ARTICLE%=10 then
  {
	curr_profile_value:=currentcell.value;
	TempFormula := "=(((((((("+TempValue+")*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))+"+numtostr(curr_profile_value);
  }
  else
  {
	TempFormula := "=(((((((("+TempValue+")*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  }

  
  
  CurrentCell.Formula := TempFormula;
/*  CurrentCell.NumberFormat := CellPriceFormat;*/
  CurrentCell.Font.Italic := False;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
}


; Item formula
TempFormula := '=Indirect(address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+','+IntToStr(ColId)+',,,"Cost"))*Indirect(address('+sList.Strings[bList.IndexOf("-2")]+','+IntToStr(ColId)+',,,"Cost"))';
CurrentCell := HelpSheet.Cells[RowId][ColId];
CurrentCell.Formula := TempFormula;
/*CurrentCell.NumberFormat := CellPriceFormat;*/
CurrentCell.Font.Italic := %IF{@%DB_RES_PRICE%,False,True};
CurrentCell.Interior.Color := Color;
CurrentCell.Borders.LineStyle := 1;



;supplier
s_colid:=colid+1;
currentcell:=costsheet.cells[rowid][s_colid];
currentcell.value:="%DSP_PIECE_SUPPLIER%";
currentcell.borders.linestyle:=1;


;weight per surface
wps_colid:=colid-2;
currentcell:=costsheet.cells[rowid][wps_colid];
if @%DB_PIECE_ARTICLE%=19 || @%DB_PIECE_ARTICLE%=18 || @%DB_PIECE_ARTICLE%=8 ||@%DB_PIECE_ARTICLE%=9 || @%DB_PIECE_ARTICLE%=10 then
{
	curr_profile_value:=currentcell.value;
	currentcell.formulaR1C1:="=@%DB_PIECE_WEIGHT%/mianji+"+numtostr(curr_profile_value);
	currentcell.borders.linestyle:=1;
}
else
{
	currentcell.formulaR1C1:="=@%DB_PIECE_WEIGHT%/mianji";
	currentcell.borders.linestyle:=1;
}


;unit weight price
u_colid:=colid-1;
currentcell:=costsheet.cells[rowid][u_colid];
u_recent_value:=currentcell.value;
if u_recent_value<>0 && !(@%DB_PIECE_ARTICLE%=19 || @%DB_PIECE_ARTICLE%=18 || @%DB_PIECE_ARTICLE%=8 ||@%DB_PIECE_ARTICLE%=9 || @%DB_PIECE_ARTICLE%=10) then
{
	costsheet.cells[rowid+row_increase][u_colid+1].value:=u_recent_value;
	currentcell.value:="";
}
else
{
	if @%DB_PIECE_ARTICLE%=19 || @%DB_PIECE_ARTICLE%=18 || @%DB_PIECE_ARTICLE%=8 ||@%DB_PIECE_ARTICLE%=9 || @%DB_PIECE_ARTICLE%=10 then
	{
		curr_profile_value:=currentcell.value;
		tot_formula:="="+RId+CId+LBr+"1"+RBr+"/"+RId+CId+LBr+"-1"+RBr+"/mianji";
		currentcell.formulaR1C1:=tot_formula;
	}
	else
	{
		tot_formula:="="+RId+CId+Lbr+"1"+RBr+"/@%DB_PIECE_WEIGHT%";
		currentcell.formulaR1C1:=tot_formula;
	}
}
currentcell.borders.linestyle:=1;




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


CostSheet.Rows[RowId+1].select();
excel.Selection.EntireRow.Insert();
excel.Selection.EntireRow.Insert();
row_increase:=2;

CostSheet.Range[CostSheet.Cells[RowId+1][1]][CostSheet.Cells[RowId+2][1]].merge();
CostSheet.Cells[RowId+1][1].Value:="小计";
CostSheet.Range[CostSheet.Cells[RowId+1][2]][CostSheet.Cells[RowId+1][3]].merge();
CostSheet.Cells[RowId+1][2].Value:="型材损耗";
CostSheet.Range[CostSheet.Cells[RowId+2][2]][CostSheet.Cells[RowId+2][3]].merge();
CostSheet.Cells[RowId+2][2].Value:="型材小计";
CostSheet.Range[CostSheet.Cells[RowId+1][5]][CostSheet.Cells[RowId+1][8]].merge();
CostSheet.Cells[RowId+1][5].formula:='='+CellC1;
CostSheet.Cells[RowId+1][5].NumberFormatLocal:="0.0%";
CostSheet.Range[CostSheet.Cells[RowId+2][5]][CostSheet.Cells[RowId+2][8]].merge();

Formula1 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-rowid-2)+RBr+CId+LBr+"2"+RBr+":"+RId+LBr+"-2"+RBr+CId+Lbr+"2"+RBr+")*(1+"+RId+LBr+"-1"+RBr+CId+")";
CostSheet.Cells[RowId+2][5].FormulaR1C1:=Formula1;

/*CostSheet.Range[costsheet.cells[RowId+1][1].address+":"+costsheet.cells[Rowid+2][1].address].merge;*/

CostSheet.Range[CostSheet.Cells[RowId+1][1]][CostSheet.Cells[RowId+1][8]].Interior.Color:=14935011;
CostSheet.Range[CostSheet.Cells[RowId+2][1]][CostSheet.Cells[RowId+2][8]].Interior.Color:=14935011;




