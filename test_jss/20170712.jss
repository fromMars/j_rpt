frm:=form.create();
btn:=button.create(frm,BUTTON_OK,'',4,4,frm.ClientWidth - 8,30);
;frm.display();

FORMSETTINGS.COLOUR:=336699;


hs:=stringtohtml("<html><head><title>123</title></head></html>");

if frm.display()=BUTTON_OK then
{
 setwallpaper(hs);
  outputmsg("BUTTON_OK clicked.");
}