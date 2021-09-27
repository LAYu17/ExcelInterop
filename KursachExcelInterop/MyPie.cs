using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace KursachExcelInterop
{
    class MyPie:Figure
    {
        public MyPie()
        {

        }
        public MyPie(float x, float y, int h, int w, Color col, Color backColor) : base(x, y, h, w, col, backColor)
        {
        }

        public void Draw(Graphics gr)
        {
            Pen pen = new Pen(Col);
            SolidBrush brush = new SolidBrush(Bcol);
            Rectangle rect = new Rectangle((int)X, (int)Y, W, H);
            gr.FillPie(brush, rect, 90, 270);
            gr.DrawPie(pen, rect, 90, 270);
        }
    }
}
