fn:=NoBackSlash(InstalledIn());
My_tempdir:=fn +'\JSS\';
fncustomlan := trim(AskFnForOpen('Select Custom language JoPPS database :', My_tempdir, 'JAdmin.db', 'db','DataBases (*.db)|*.db'));
lenfncustom:=strlen(fncustomlan);
fncustom := strleft(fncustomlan,lenfncustom-9);
fns:=fn+'\LAN\';
db_newlan := dbtable.create();
db_newlan.DataBaseName := fns;
db_newlan.tableName := 'Jopps.db';
db_newlan.openexclusive();

/*CHECK NEW LANGUAGE-DATABASE JOPPS*/




db_newlan.Active :=True;



db_newlan.First();

while !(db_newlan.eof) do
{

IDCTRL:=db_newlan.field['CTRL_ID'];
IDFORM:=db_newlan.field['FORM_ID'];
S_IDCTRL:=trim(numtostr(IDCTRL,8,0));
S_IDFORM:=trim(numtostr(IDFORM,8,0));


/*CHECK CUSTOM LANGUAGE-DATABASE JOPPS*/

db_customlan:= dbquery.create();
db_customlan.DataBaseName :=fncustom;
db_customlan.open();

db_customlan.Sql:= 'SELECT FORM_ID,CTRL_ID,LAN_USER,HIN_USER,TOP_USER FROM JOPPS.db where FORM_ID='+S_IDFORM+' and CTRL_ID='+S_IDCTRL+"'";


db_customlan.Active :=True;


db_customlan.First();

while !(db_customlan.eof) && db_customlan.recordcount>0 do
{
db_customlan.edit();
db_newlan.edit();
canreadwrite:=db_newlan.canmodify;
db_newlan.field['LAN_USER']:=db_customlan.field['LAN_USER'];
db_newlan.field['HIN_USER']:=db_customlan.field['HIN_USER'];
db_newlan.field['TOP_USER']:=db_customlan.field['TOP_USER'];
db_newlan.post();
db_customlan.post();

db_customlan.next();
};

db_customlan.Clearfields();
db_customlan.Close();
db_customlan.Free();

db_newlan.next();
};

db_newlan.Clearfields();
db_newlan.Close();
db_newlan.Free();

