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


/*curr_assembly:=getcurrentproject().projectdata.children[cnt-1];*/
curr_assembly:=getcurrentproject().projectdata.currentassembly;
assembly_cnt:=getcurrentproject().projectdata.childcount;
i_cnt:=0;
while i_cnt<assembly_cnt do
{
    curr_assembly:=getcurrentproject().projectdata.children[i_cnt];
    if curr_assembly.code="%ASSEMBLY_TEXT%" then
        break;
    i_cnt:=i_cnt+1;
}

;assembly name
curr_name:=curr_assembly.code;
/*costsheet.range["chuanghao"].value:="窗号："+"@%DB_PIECE_ASSEMBLY%";*/
costsheet.range["chuanghao"].value:=curr_assembly.code;

;surface
f_cnt:=0;
curr_frame:=curr_assembly.children[0];
frame_cnt:=curr_assembly.childcount;
a_mianji:=0;
while f_cnt<frame_cnt do
{
    curr_frame:=curr_assembly.children[f_cnt];
    f_width:=curr_frame.width;
    f_height:=curr_frame.height;
    f_mianji:=f_width*f_height;
    a_mianji:=a_mianji+f_mianji;
    f_cnt:=f_cnt+1;
}
/*
curr_width:=curr_assembly.width;
curr_height:=curr_assembly.height;*/
curr_surface:=a_mianji/1000000;
costsheet.range["mianji"].value:=curr_surface;
costsheet.range["mianji"].HorizontalAlignment:=-4131;
costsheet.range["mianji"].offset[0][-1].HorizontalAlignment:=-4152;

a_fee_row:=0;

/*used to calculate A*/
RowId_0:=0;
RowId_1:=0;
RowId_2:=0;
RowId_A:=0;

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
  currentcell.NumberFormat:=CellCostFormat;
  CurrentCell1 := CostSheet.Cells[RowId][ColId-1];
  if CurrentCell1.Value=0 then
    CurrentCell1.Value := TempValue;
  currentcell1.NumberFormat:=CellCostFormat;
  CurrentCell0 := CostSheet.Cells[RowId][ColId-2];
  CurrentCell0.Value := TempValue;
  /*CurrentCell.NumberFormat := CellPriceFormat;*/
  CurrentCell.Font.Italic := True;
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
  i := i + 1;
};

/*calculate follow artikels, recent_profile_value-recent TempValue[string],tmp_tmp_value-current TempValue[string]*/
recent_profile_value:="0";
tmp_tmp_value:="0";

a_linked:=strings.create();

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


a_link:="";
z_pg:=pricegroups.create();
z_pg.code.group:="A";
z_pg.code.block:=@%DB_COST_ARTICLE%;
if z_pg.find() then
{
	a_link:=z_pg.link;
    if a_link<>"" && a_linked.indexof(a_link)=-1 then
        a_linked.add(a_link);
}
else
{
	msgbox("no article block "+inttostr(z_pg.code.block)+" found!");
}



; Item price
/*RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%"+"@%DB_COST_RATION%"+"@%DB_COST_FACTOR%"+"@%DB_COST_RATIO%")]);*/
if a_link<>"" then
{
	RowId  := StrToNum(cList.Strings[bList.IndexOf(a_link+"@%DB_PIECE_LOSSTYPE%")]);
}
else
{
	RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_PIECE_ARTICLE%"+"@%DB_PIECE_LOSSTYPE%")]);
}


/*
if a_link="" then
	RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_PIECE_ARTICLE%"+"@%DB_PIECE_LOSSTYPE%")]);
else
	RowId  := StrToNum(cList.Strings[bList.IndexOf(a_link+"@%DB_PIECE_LOSSTYPE%")]);*/



if recent_rowid=-1 || recent_rowid>rowid then
	recent_rowid:=rowid;

CellCT := 'Indirect("Cost!"&address('+sList.Strings[cList.IndexOf(IntToStr(RowId))]+","+IntToStr(ColCT)+"))/"+pList.Strings[cList.IndexOf(IntToStr(RowId))];
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
  if a_link<>"" || a_linked.indexof("@%DB_COST_ARTICLE%")<>-1 then
  {
	tmp_tmp_value:=tempvalue;
	curr_profile_value:="((((((((("+recent_profile_value+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
	TempFormula := "=((((((((("+TempValue+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))+"+curr_profile_value;
  }
  else
  {
	TempFormula := "=((((((((("+TempValue+")*("+CellCT+"))*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
  }
  
  
  CurrentCell.Formula := TempFormula;
/*  CurrentCell.NumberFormat := CellPriceFormat;*/
/*  CurrentCell.Font.Italic := False;*/
  CurrentCell.Interior.Color := Color;
  CurrentCell.Borders.LineStyle := 1;
}
else
{
  TempValue   := StrReplace("@%DB_PIECE_PRICE%/%ASSEMBLYCOUNT%",".","%DECIMALSEP%");
  CurrentCell := CostSheet.Cells[RowId][ColId];

  if a_link<>"" || a_linked.indexof("@%DB_COST_ARTICLE%")<>-1 then
  {
	tmp_tmp_value:=tempvalue;
	curr_profile_value:="(((((((("+recent_profile_value+")*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))";
	TempFormula := "=(((((((("+TempValue+")*(1+"+CellC1+"))*(1-"+CellC2+"))*"+CellC7+")*"+CellC3+")*(1+"+CellC6+"))*(1+"+CellC4+"))*(1-"+CellC5+"))+"+curr_profile_value;
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



;unit name
un_colid:=3;
currentcell:=costsheet.cells[rowid][un_colid];
currentcell.value:="Kg";
currentcell.HorizontalAlignment:=-4108;


;supplier
s_colid:=colid+1;
currentcell:=costsheet.cells[rowid][s_colid];
currentcell.value:="%DSP_PIECE_SUPPLIER%";
currentcell.borders.linestyle:=1;


;weight per surface
wps_colid:=colid-2;
currentcell:=costsheet.cells[rowid][wps_colid];
if a_link<>"" || a_linked.indexof("@%DB_COST_ARTICLE%")<>-1 then
{
	curr_profile_value:=currentcell.value;
	/*currentcell.formulaR1C1:="=@%DB_PIECE_WEIGHT%/mianji+"+numtostr(curr_profile_value);*/
    currentcell.formulaR1C1:="=@%DB_PIECE_WEIGHT%/%ASSEMBLYCOUNT%+"+numtostr(curr_profile_value);
	currentcell.borders.linestyle:=1;
}
else
{
	/*currentcell.formulaR1C1:="=@%DB_PIECE_WEIGHT%/mianji";*/
    currentcell.formulaR1C1:="=@%DB_PIECE_WEIGHT%/%ASSEMBLYCOUNT%";
	currentcell.borders.linestyle:=1;
}


;unit weight price
u_colid:=colid-1;
currentcell:=costsheet.cells[rowid][u_colid];
u_recent_value:=currentcell.value;
if u_recent_value<>0 && a_link="" && a_linked.indexof("@%DB_COST_ARTICLE%")=-1 then
{
	costsheet.cells[rowid][u_colid+1].value:=u_recent_value;
	currentcell.value:="";
}
else
{
	if a_link<>"" || a_linked.indexof("@%DB_COST_ARTICLE%")<>-1 then
	{
		curr_profile_value:=currentcell.value;
		/*tot_formula:="="+RId+CId+LBr+"1"+RBr+"/"+RId+CId+LBr+"-1"+RBr+"/mianji";*/
        tot_formula:="="+RId+CId+LBr+"1"+RBr+"/"+RId+CId+LBr+"-1"+RBr;
		currentcell.formulaR1C1:=tot_formula;
	}
	else
	{
		tot_formula:="="+RId+CId+Lbr+"1"+RBr+"/(@%DB_PIECE_WEIGHT%/%ASSEMBLYCOUNT%)";
		currentcell.formulaR1C1:=tot_formula;
	}
}
currentcell.borders.linestyle:=1;

/*calculate follow artikels*/
recent_profile_value:=tmp_tmp_value;
tmp_tmp_value:="0";

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
costsheet.cells[rowid+1][1].VerticalAlignment:=-4108;
costsheet.cells[rowid+1][1].HorizontalAlignment:=-4108;
CostSheet.Range[CostSheet.Cells[RowId+1][2]][CostSheet.Cells[RowId+1][3]].merge();
CostSheet.Cells[RowId+1][2].Value:="型材损耗";
CostSheet.Range[CostSheet.Cells[RowId+2][2]][CostSheet.Cells[RowId+2][3]].merge();
CostSheet.Cells[RowId+2][2].Value:="型材小计";
CostSheet.Range[CostSheet.Cells[RowId+1][5]][CostSheet.Cells[RowId+1][7]].merge();

/*CostSheet.Cells[RowId+1][5].formula:='='+CellC1;*/
CostSheet.Cells[RowId+1][5].value:=0;

CostSheet.Cells[RowId+1][5].NumberFormatLocal:="0.0%";
CostSheet.Range[CostSheet.Cells[RowId+2][5]][CostSheet.Cells[RowId+2][7]].merge();
costsheet.cells[rowid+2][5].NumberFormat:=CellCostFormat;

Formula1 := "="+SumFormulaText+"("+RId+LBr+IntToStr(recent_rowid-rowid-2)+RBr+CId+LBr+"2"+RBr+":"+RId+LBr+"-2"+RBr+CId+Lbr+"2"+RBr+")*(1+"+RId+LBr+"-1"+RBr+CId+")";
CostSheet.Cells[RowId+2][5].FormulaR1C1:=Formula1;

/*CostSheet.Range[costsheet.cells[RowId+1][1].address+":"+costsheet.cells[Rowid+2][1].address].merge;*/

CostSheet.Range[CostSheet.Cells[RowId+1][1]][CostSheet.Cells[RowId+1][8]].Interior.Color:=14935011;
CostSheet.Range[CostSheet.Cells[RowId+2][1]][CostSheet.Cells[RowId+2][8]].Interior.Color:=14935011;

Rowid_0:=RowId+2;


