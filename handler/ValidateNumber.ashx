<%@ WebHandler Language="C#" Class="ValidateNumber" %>

using System;
using System.Web;
using System.Drawing;
using System.Web.SessionState;

public class ValidateNumber : IHttpHandler, IRequiresSessionState
{

    public void ProcessRequest(HttpContext context)
    {

        int NumCount = 4;//預設產生5位亂數

        //取得亂數
        string str_ValidateCode = this.GetRandomNumberString(NumCount);
        /*用於驗證的Session*/
        context.Session["ValidateNumber"] = str_ValidateCode;

        //取得圖片物件
        System.Drawing.Image image = this.CreateCheckCodeImage(str_ValidateCode);
        System.IO.MemoryStream ms = new System.IO.MemoryStream();
        image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
        /*輸出圖片*/
        context.Response.Clear();
        context.Response.ContentType = "image/jpeg";
        context.Response.BinaryWrite(ms.ToArray());
        ms.Close();

    }

    #region 產生亂數
    private string GetRandomNumberString(int int_NumberLength)
    {
        //純數字
        //System.Text.StringBuilder str_Number = new System.Text.StringBuilder();//字串儲存器
        //Random rand = new Random(Guid.NewGuid().GetHashCode());//亂數物件

        //for (int i = 1; i <= int_NumberLength; i++)
        //{
        //    str_Number.Append(rand.Next(0, 10).ToString());//產生0~9的亂數
        //}
        //return str_Number.ToString();

        //英+數
        string allChar = "0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z";
        string[] allCharArray = allChar.Split(',');   //以,做為區隔並寫入陣列
        string randomCode = "";
        int temp = -1;

        Random rand = new Random(Guid.NewGuid().GetHashCode());   //使用GUID產生亂數
        for (int i = 0; i < int_NumberLength; i++)
        {
            if (temp != -1)
            {
                rand = new Random(Guid.NewGuid().GetHashCode());    //使用GUID產生亂數
            }
            int t = rand.Next(36);      //共36個(10數字+26字母)
            if (temp != -1 && temp == t)
            {
                return GetRandomNumberString(int_NumberLength);
            }
            temp = t;
            randomCode += allCharArray[t];
        }
        return randomCode;
    }
    #endregion

    #region 產生圖片
    private System.Drawing.Image CreateCheckCodeImage(string checkCode)
    {
        System.Drawing.Bitmap image = new System.Drawing.Bitmap((checkCode.Length * 23), 35);//產生圖片，寬23*位數，高40像素
        System.Drawing.Graphics g = Graphics.FromImage(image);

        //生成隨機生成器
        Random random = new Random(Guid.NewGuid().GetHashCode());
        int int_Red = 0;
        int int_Green = 0;
        int int_Blue = 0;
        int_Red = random.Next(256);//產生0~255
        int_Green = random.Next(256);//產生0~255
        int_Blue = (int_Red + int_Green > 400 ? 0 : 400 - int_Red - int_Green);
        int_Blue = (int_Blue > 255 ? 255 : int_Blue);

        //清空圖片背景色
        //g.Clear(Color.FromArgb(int_Red, int_Green, int_Blue));
        g.Clear(Color.White);

        //畫圖片的背景噪音線
        for (int i = 0; i <= 24; i++)
        {
            int x1 = random.Next(image.Width);
            int x2 = random.Next(image.Width);
            int y1 = random.Next(image.Height);
            int y2 = random.Next(image.Height);

            g.DrawLine(new Pen(Color.Silver), x1, y1, x2, y2);

            g.DrawEllipse(new Pen(Color.Teal), new System.Drawing.Rectangle(x1, y1, x2, y2));
        }

        System.Drawing.Drawing2D.LinearGradientBrush brush = new System.Drawing.Drawing2D.LinearGradientBrush(new Rectangle(0, 0, image.Width, image.Height), Color.Blue, Color.DarkRed, 1.2F, true);

        Font font = new System.Drawing.Font("Arial", 20, (System.Drawing.FontStyle.Bold));
        g.DrawString(checkCode, font, brush, 2, 2);

        for (int i = 0; i <= 99; i++)
        {
            //畫圖片的前景噪音點
            int x = random.Next(image.Width);
            int y = random.Next(image.Height);

            image.SetPixel(x, y, Color.FromArgb(random.Next()));
        }

        //畫圖片的邊框線
        g.DrawRectangle(new Pen(Color.Silver), 0, 0, image.Width - 1, image.Height - 1);

        return image;
    }
    #endregion


    public bool IsReusable {
        get {
            return false;
        }
    }

}