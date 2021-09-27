using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
namespace KursachExcelInterop
{
    class MyRectangle:Figure

    {
        public MyRectangle()
        {

        }

        public MyRectangle(double x, double y, int h, int w, Color col, Color backColor) : base(x, y, h, w, col, backColor)
        {
        }

        public void Draw(Graphics gr)
        {
            Pen pen = new Pen(Col);
            SolidBrush brush = new SolidBrush(Bcol);
            gr.FillRectangle(brush, (float)X, (float)Y, W, H);
            gr.DrawRectangle(pen, (float)X, (float)Y, W, H);
        }
    }
}
