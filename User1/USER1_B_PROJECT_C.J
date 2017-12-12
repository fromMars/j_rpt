/* USER1_B_PROJECT_C.J
 * get handling tab values from charge table
 * correct quantity per assembly */



; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)
; 


%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail   b_project_C.j
; 

RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%")]);


;单樘用量
wps_colid:=colid-1;
currentcell:=costsheet.cells[rowid][wps_colid];
currentcell0:=costsheet.cells[rowid][wps_colid-3];
/*currentcell0.HorizontalAlignment:=-4108; useless*/
if @%DB_CHARGE_QUANTITY%<>0 then                                /*输出单樘用量*/
{
    currentcell.formulaR1C1:="=@%DB_CHARGE_QUANTITY%";
}
else
	currentcell.value:=0;
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


