/* USER1_E_PROJECT.J
 * Summary Table
 *                      */

helpsheet:=template.worksheets['Help'];
sheet_count:=template.worksheets.count-3;
helpsheet.cells[1][1].value:="%PROJECT_TEXT%";

c_cnt:=1;
c_name:=2;
c_type:=3;
c_size_w:=4;
c_size_h:=5;
c_qty:=6;
c_area:=7;
c_price_per_assembly:=8;
c_price_per_area:=9;
c_price_summary:=10;

s0:=1;
while s0<=sheet_count do
{
    tmp:=template.worksheets[s0];
    s:=s0+3;
    
    v_cnt:='=row()-3';
    v_name:=tmp.name;
    v_type:='';
    v_size_w:=0;
    v_size_h:=0;
    v_qty:=0;
    v_area:='='+v_name+'!mianji';
    v_price_per_area:='='+v_name+'!danjia';
    v_price_per_assembly:='='+v_name+'!danjia*'+v_name+'!mianji';
    v_price_summary:='=RC[-2]*RC[-4]';
    
    helpsheet.cells[s][c_cnt].formula:=v_cnt;
    helpsheet.cells[s][c_name].value:=substr(v_name,6,strlen(v_name));
    helpsheet.cells[s][c_type].value:=v_type;
    /*helpsheet.cells[s][c_size_w].value:=v_size_w;
    helpsheet.cells[s][c_size_h].value:=v_size_h;
    helpsheet.cells[s][c_qty].value:=v_qty;*/
    helpsheet.cells[s][c_area].value:=v_area;
    helpsheet.cells[s][c_price_per_assembly].formula:=v_price_per_assembly;
    helpsheet.cells[s][c_price_per_area].formula:=v_price_per_area;
    helpsheet.cells[s][c_price_summary].formular1c1:=v_price_summary;
    
    s0:=s0+1;
    
}

helpsheet.usedrange.borders.linestyle:=1;
helpsheet.columns.Autofit;




