
doc=app.activeDocument;
pageObj=doc.pages.item(0);
txtObj=pageObj.textFrames.item(0);

var temp=doc.pages.length;

for(var i= 0; i<temp;i++)
{
    if(doc.pages[i].textFrames.item(0)!=null)
    {
            if(doc.pages[i].textFrames[0].contents=="")
            {
                doc.pages[i].remove();
                i=i-1;
                temp--;
            }
    }
    else
    {
      doc.pages[i].remove();
      i=i-1;
      temp--;
    }

}