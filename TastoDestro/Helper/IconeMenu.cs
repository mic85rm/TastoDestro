using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using stdole;
public class IconeMenu : AxHost
{
    public IconeMenu() : base("59EE46BA-677D-4d20-BF10-8D8067CB8B33")
    {
    }
    public static new stdole.StdPicture GetIPictureDispFromPicture(Image image)
    {
        return (StdPicture)AxHost.GetIPictureDispFromPicture(image);
    }


    public static Image LoadBase64(string base64)
    {
        byte[] bytes = Convert.FromBase64String(base64);
        Image image;
        using (MemoryStream ms = new MemoryStream(bytes))
        {
            image = Image.FromStream(ms);
        }
        return image;
    }

}