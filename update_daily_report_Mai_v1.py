# -*- coding: utf-8 -*-
"""
Created on Mon Aug  9 12:56:40 2021

@author: ngothimai
"""


from pathlib import Path
import sys

import os
import win32com.client as client
from PIL import ImageGrab
import datetime as datet
from datetime import datetime
from dateutil.relativedelta import relativedelta
import time
win32c = client.constants
now = datetime.now()
last_month = now+relativedelta(months=-1)
def run_sendmail(f_path: Path,  f_name: str, subject: str, to: str, cc: str ,bcc:str) -> list:

    filename = f_path / f_name
    # filename1 = f_path1/f_name1

    # create excel object
    excel = client.gencache.EnsureDispatch('Excel.Application')

    # excel can be vsisible or not
    excel.Visible = True
    
    # try except for file / path
    try:
        wb = excel.Workbooks.Open(filename)
        time.sleep(2)
       # wb1 = excel.Workbooks.Open(filename1,UpdateLinks=False)
        # time.sleep(2)
    except error as e:
        if e.excepinfo[5] == -2146827284:
            print(f'Failed to open spreadsheet.  Invalid filename or location: {filename}')
            print(f'Failed to open spreadsheet.  Invalid filename or location: {filename1}')
            sys.exit(1)
        pass
  
    outlook = client.Dispatch('Outlook.Application')
    # create a message
    message = outlook.CreateItem(0)
    # copy Excel table as Image
    wb.Sheets['funnel_M_N'].Range('A7:I24').CopyPicture(Format=win32c.xlBitmap)
    
    # startup and instance of outlook
    
    # set the message properties and content
    img = ImageGrab.grabclipboard()
    image_path = sys.path[0]+'\\summary_0.png'
    img.save(image_path)
    message.Attachments.Add(Source=image_path)

    wb.Sheets['Funnel_Month N-1'].Range('A7:D23').CopyPicture(Format=win32c.xlBitmap)
    img = ImageGrab.grabclipboard()
    image_path = sys.path[0]+'\\summary_1.png'
    img.save(image_path)
    message.Attachments.Add(Source=image_path)
    
    wb.Sheets['Conversion rate'].Range('A7:J20').CopyPicture(Format=win32c.xlBitmap)
    img = ImageGrab.grabclipboard()
    image_path = sys.path[0]+'\\summary_2.png'
    img.save(image_path)
    message.Attachments.Add(Source=image_path)

    wb.Sheets['Conversion rate'].Range('A35:J41').CopyPicture(Format=win32c.xlBitmap)
    img = ImageGrab.grabclipboard()
    image_path = sys.path[0]+'\\summary_3.png'
    img.save(image_path)
    message.Attachments.Add(Source=image_path)
    # wb.Sheets['Summary'].Range('A208:Q219').CopyPicture(Format=win32c.xlBitmap)
    # img = ImageGrab.grabclipboard()
    # image_path = sys.path[0]+'\\summary_3.png'
    # img.save(image_path)
    # message.Attachments.Add(Source=image_path)
    
    # copy Excel table as Image
    """
     wb.Sheets['Section1'].Range('B18:AH44').CopyPicture(Format=win32c.xlBitmap)
    # set the message properties and content
    # img = ImageGrab.grabclipboard()
    image_path = sys.path[0]+'\\section1.png'
    img.save(image_path)
    message.Attachments.Add(Source=image_path)
    
    wb.Sheets['Section2'].Range('A18:AG32').CopyPicture(Format=win32c.xlBitmap)
    wb.Close(False)    
    img = ImageGrab.grabclipboard()
    image_path = sys.path[0]+'\\section2.png'
    img.save(image_path)
    message.Attachments.Add(Source=image_path)
    
    wb1.Sheets['Dashboard'].Range('AI1:AX19').CopyPicture(Format=win32c.xlBitmap)
    wb1.Save()
    wb1.Close(False)
    excel.Quit()
    img = ImageGrab.grabclipboard()
    image_path = sys.path[0]+'\\dashboard.png'
    img.save(image_path)
    message.Attachments.Add(Source=image_path)
  """
    #signatureimage1 = "D:/SOFT/WPy64-3740/settings/image_ocb/ocb.png"
    #signatureimage2 = "D:/SOFT/WPy64-3740/settings/image_ocb/icon_fb.png"
    #signatureimage3 = "D:/SOFT/WPy64-3740/settings/image_ocb/icon_instagram.png"
    #signatureimage4 = "D:/SOFT/WPy64-3740/settings/image_ocb/icon_linkedin.png"
    #signatureimage5 = "D:/SOFT/WPy64-3740/settings/image_ocb/icon_yt.png"
    #signatureimage6 = "D:/SOFT/WPy64-3740/settings/image_ocb/icon_zalo.png"
    signatureimage7 = "D:/SOFT/WPy64-3740/settings/image_ocb/img_icon.png"
    #attachment1 = message.Attachments.Add(Source=signatureimage1)
    #attachment2 = message.Attachments.Add(Source=signatureimage2)
    #attachment3 = message.Attachments.Add(Source=signatureimage3)
    #attachment4 = message.Attachments.Add(Source=signatureimage4)
    #attachment5 = message.Attachments.Add(Source=signatureimage5)
    #attachment6 = message.Attachments.Add(Source=signatureimage6)
    attachment7 = message.Attachments.Add(Source=signatureimage7)
    #attachment1.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "ocb")
    #attachment2.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "fb")
    #attachment3.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "insta")
    #attachment4.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "linkedin")
    #attachment5.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "yt")
    #attachment6.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "zalo")
    attachment7.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "icon")
    
    html_body = """
   
    <div>
          Dear anh Quí,
          <br>
          <br>
          Em gửi anh Funnel Topup Covid Daily Report """ + now.strftime("%d %b, %Y") + """

    </div>
    <div>
         <h3 style="color:red">+ Funnel daily </h3>
         <p  style="font-weight: bold">Month import: """ + now.strftime("%Y-%m") + """ </p>
        <img src=summary_0.png>
    <br>
    <br>
        <p style="font-weight: bold">Month import: """ + last_month.strftime("%Y-%m") + """ </p>
        <img src=summary_1.png>
    </div>
    <div>
        <h3 style="color:red">+ Conversion rate daily </h3>
        <p  style="font-weight: bold">Month import: """ + now.strftime("%Y-%m") + """ </p>
        <img src=summary_2.png>
    </div>
    <br>
    <br>
    <div>
        <p  style="font-weight: bold">Month import: """ + last_month.strftime("%Y-%m") + """ </p>
        <img src=summary_3.png>
    </div>
    <br>
    <div>

       
        <p> Báo cáo chi tiết anh vui lòng xem ở link: <a href="W:/OfficeShare/15.Product development/01. WEEKLY REPORT TO GĐK/Lead Pool and Funnel/Report/Report Leadgen RF file - store for ppt_TOPUPCOVID.xlsx">W:/OfficeShare/15.Product development/01. WEEKLY REPORT TO GĐK/Lead Pool and Funnel/Report/Report Leadgen RF file - store for ppt_TOPUPCOVID.xlsx </a> </p>
        
    </div>
     <br> 
    <div>
    <h4> Regards, </h4>
    </div>
    <div>
    <div>
    NGÔ THỊ MAI
    <br>
    Chuyên viên Quản lý danh mục
    <br>
    Phòng Phát triển sản phẩm
    <br>
    Mobile: 0947.068.476
    <br>
    
    </div>

    </div>
    <div>
        <p>Địa chỉ: Tầng 1- Tòa nhà Blue Square- 91 Phạm Văn Hai- Phường 3- Quận Tân Bình- TP HCM 
        <br>
        Tel: (+84-8) 3622 0139 
        <br>
        Email: ngothimai@m-ocb.com.vn |  Website: www.com-b.vn 
        </p>
    </div>
    <img src="cid:icon">
    """
    message.HTMLBody = html_body
    
    message.To = to 
    message.CC = cc
    message.BCC= bcc
    message.Subject = subject

    # display the message to review
    message.Display()
    
    # save or send the message
    
    #message.Send() 
if __name__ == "__main__":
    # file path
    f_path = Path(r'W:/OfficeShare/15.Product development/01. WEEKLY REPORT TO GĐK/Lead Pool and Funnel/Report') # file in current working directory
    # f_path1=Path(r'W:\OfficeShare\15.Product development\98. Report\28. Vol Daily Tracking')
    # excel filename
    
    f_name = 'Report Leadgen RF file - store for ppt_TOPUPCOVID.xlsx'
    # f_name1='Daily tracking v2.xlsx'
    f_tmp = f_path/f_name
    modified = os.path.getmtime(f_tmp)
    year,month,day,hour,minute,second=time.localtime(modified)[:-3]
    d= '%02d-%02d-%02d'%(year,month,day)
    n= datet.datetime.today().strftime('%Y-%m-%d')
    to = "phuongbaoqui@m-ocb.com.vn \
            ;doduongbaouyen@m-ocb.com.vn \
            ;dongvantinh@m-ocb.com.vn \
            ;nguyenptanguyet@m-ocb.com.vn \
            ;nguyentuanthanh@m-ocb.com.vn \
            ;vonhutthu@m-ocb.com.vn"
    cc= "hoangcaodat@m-ocb.com.vn;nguyphannam@m-ocb.com.vn;hoangminhhaingan@m-ocb.com.vn;nguyenthinhung1@m-ocb.com.vn"
    
    #if datet.datetime.today().weekday()!=6:
    if datet.datetime.today().weekday()!=6 \
        and ((now.replace(day=1)-relativedelta(months=1)).weekday()!=6 or now.day!=2) \
        and d==n: #now.day!=2 and now.day!=3 
        # function calls
        #run_sendmail(f_path,f_path1, f_name,f_name1, subject= "Leadgen Daily Report " + now.strftime("%d %b, %Y"), to="nguyenthanhphu@m-ocb.com.vn", cc="KHDCkhoitaonguonkh@m-ocb.com.vn;KHDCPTSP@m-ocb.com.vn;KHDCKHDL@m-ocb.com.vn;khdcphantichkinhdoanh@m-ocb.com.vn;KHDCDigitalLending@m-ocb.com.vn;hoanganhthao@m-ocb.com.vn;tranthianhminh@m-ocb.com.vn; nguyenhoan@m-ocb.com.vn;phuongbaoqui@m-ocb.com.vn;nguyenanhhai@m-ocb.com.vn;hoangcaodat@m-ocb.com.vn;lephamthuytien@m-ocb.com.vn;buiduykhanh1@m-ocb.com.vn", bcc="")
        #run_sendmail(f_path, f_name, subject= "Funel Topup Covid Daily Report " + now.strftime("%d %b, %Y"), to="phuongbaoqui@m-ocb.com.vn;doduongbaouyen@m-ocb.com.vn;dongvantinh@m-ocb.com.vn;nguyenptanguyet@m-ocb.com.vn;nguyentuanthanh@m-ocb.com.vn;vonhutthu@m-ocb.com.vn",cc="hoangcaodat@m-ocb.com.vn;nguyphannam@m-ocb.com.vn;hoangminhhaingan@m-ocb.com.vn;nguyenthinhung1@m-ocb.com.vn",bcc="")
        run_sendmail(f_path, f_name, subject= "Funel Topup Covid Daily Report " + now.strftime("%d %b, %Y"), to=to,cc=cc,bcc="")

    else:
        print('cannot send mail in sunday!')