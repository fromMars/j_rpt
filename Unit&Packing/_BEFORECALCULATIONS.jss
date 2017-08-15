CurPro := GetCurrentProject();
IF CurPro = Nil THEN halt;
d:=CurPro.setup.profileprice;
/*f_name:="test"+getparam("PROJECT_TEXT");*/
if getparam("ORDER_LIST_PER_SUPPLIER_PRICE")="1" then
{
	if d=1 then
	{
		msgbox("change to 3");
		setparam("ORDER_LIST_PER_SUPPLIER_PRICE","3");
		if fo:=openfile("test"+getparam("PROJECT_TEXT")) then
		{
		    closefile(fo);
		    deletefile("test"+getparam("PROJECT_TEXT"));
		}
		fc:=createfile("test"+getparam("PROJECT_TEXT"));
		writestr(fc,"3");
		closefile(fc);
					
										
	}
	else if d=0 then
	{
		msgbox("change to 1");
		/*setparam("ORDER_LIST_PER_SUPPLIER_PRICE","3");*/
		if fo:=openfile("test"+getparam("PROJECT_TEXT")) then
		{
		    closefile(fo);
		    deletefile("test"+getparam("PROJECT_TEXT"));
		}
		fc:=createfile("test"+getparam("PROJECT_TEXT"));
		writestr(fc,"1");
		closefile(fc);
	}
}
/*
fc:=createfile("test"+getparam("PROJECT_TEXT"));
rs:=readstr(fc);
msgbox("end:"+rs);
closefile(fc);*/
msgbox("end:"+getparam("ORDER_LIST_PER_SUPPLIER_PRICE"));



/*msgbox("head:"+getparam("ORDER_LIST_PER_SUPPLIER_PRICE"));
CurPro := GetCurrentProject();
IF CurPro = Nil THEN halt;
d:=CurPro.setup.profileprice;
if getparam("ORDER_LIST_PER_SUPPLIER_PRICE")="1" then
{
	msgbox("in 1");
	if d=1 then
	{
		msgbox("change to 3");
		setparam("ORDER_LIST_PER_SUPPLIER_PRICE","3");
	}
}
else if getparam("ORDER_LIST_PER_SUPPLIER_PRICE")="3" then
{
	msgbox("in 3");
	if d=0 then
	{
		msgbox("change to 1");
		setparam("ORDER_LIST_PER_SUPPLIER_PRICE","1");
	}
}
msgbox("end:"+getparam("ORDER_LIST_PER_SUPPLIER_PRICE"));
*/