/* USER1_B_ASSEMBLY_1_A.J
 * correct quantity per assembly */



; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%)
; 


%% detail
; ******************************Estim Excel************************************
; *****************************************************************************
; %NAME% (%BATCH%) - Detail   b_assembly_1_A.j
; 



; Item price
RowId  := StrToNum(cList.Strings[bList.IndexOf("@%DB_COST_ARTICLE%"+"@%DB_COST_LOSSTYPE%")]);
if recent_rowid=-1 || recent_rowid>(rowid+row_increase) then
	recent_rowid:=rowid+row_increase;



/*单樘用量*/
wps_colid:=colid-2;
currentcell:=costsheet.cells[rowid+row_increase][wps_colid];
currentcell0:=costsheet.cells[rowid+row_increase][wps_colid-2];
currentcell0.HorizontalAlignment:=-4108;
if @%DB_ATTRIB_NO%<>0 || @%DB_ATTRIB_CFLENGTH%<>0 then
{
    currentcell.formulaR1C1:="=@%DB_ATTRIB_NO%/%ASSEMBLYCOUNT%";
    currentcell0.value:="件";
    /*价格块35-胶条*/
    if @%DB_ATTRIB_ARTICLE%=35 then
    {
        currentcell.formulaR1C1:="=@%DB_ATTRIB_CFLENGTH%/%ASSEMBLYCOUNT%";
        currentcell0.value:="m";
    }   
    
}
else
	currentcell.value:="";
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


