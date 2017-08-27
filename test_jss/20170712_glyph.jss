frm:=form.create("test",400,500);
btn:=button.create(frm,button_ok,"ok",300,350,50,39);
btn1:=button.create(frm,button_cancel,"cancel",300,400,50,39);

gg:=glyph.create(frm,glyph_YES,"",10,350,50,39);
lb:=label.create(frm,inttostr(frm.height)+", "+inttostr(frm.width),5,5);

gg.load("test.bmp");

outputmsg(inttostr(frm.height)+", "+inttostr(frm.width));
while frm.display()<>button_cancel do
{
 lb.caption:=numtostr(frm.winhandle);
}