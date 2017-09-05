/* Isolator Selector
 * Applied to FramePart System ES70A, ES78, ES101 */

p:=getcurrentproject();
pc_cnt:=p.projectdata.childcount;
i:=0;
while i<pc_cnt do
{
    a:=p.projectdata.children[i];
    ac_cnt:=a.childcount;
    j:=0;
    outputmsg(inttostr(ac_cnt));
    while j<ac_cnt do
    {
        ac:=a.children[j];
        ac_system:=ac.model.system;
        if ac.isframepart then
        {
            if ac_system="ES70A" || ac_system="ES78" || ac_system="ES101" then
                ac.isolator:=2;
            else
                ac.isolator:=1;
        }
        j:=j+1;
    }
    i:=i+1;
}
