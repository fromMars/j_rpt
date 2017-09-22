a:=getcurrentassembly();
cnt:=a.childcount;
while cnt>0 do
{
    cnt:=cnt-1;
    c:=a.children[cnt];
    if c.isframepart then                               /* framepart */
    {
        b_cnt:=c.childcount;
        while b_cnt>0 do
        {
            b_cnt:=b_cnt-1;
            d:=c.children[b_cnt];
            if d.isframeelement then                        /* frameelement */
            {
                x:=d.profile.system;
                if d.profile.system="ES70A" then
                {
                    msgbox("find ES70A");
                }
            }
            else if d.isframeopening then                   /* frameopening */
            {
                m_cnt:=d.childcount;
                while m_cnt>0 do
                {
                    m_cnt:=m_cnt-1;
                    m:=d.children[m_cnt];
                    if m.isventpart then                        /* ventpart */
                    {
                        p_cnt:=m.childcount;
                        while p_cnt>0 do
                        {
                            p_cnt:=p_cnt-1;
                            p:=m.children[p_cnt];
                            if p.isventelement then                 /* ventelement */
                            {
                                y:=p.profile.system;
                                if p.profile.system="ES70A" then
                                {
                                    msgbox("find ES70A");
                                }
                            }
                        }
                    }
                }
            }
        }
        /*msgbox(c.isolator,", ",c.addon[0].profile.system,",",cnt);*/
    }
}
