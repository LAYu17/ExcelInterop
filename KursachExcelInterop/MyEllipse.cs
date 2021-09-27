using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace KursachExcelInterop
{
    class MyEllipse:Figure
    {
        private int activePoint = 1;
        public MyEllipse()
        {

        }
        public MyEllipse(double x, double y, int h, int w, Color col, Color backColor) : base(x, y, h, w, col, backColor)
        {
        }
        public bool Shot(float xx, float yy)
        {
            activePoint = -1; // никуда не попал
            //попал в тело
            if ((xx >= X && xx <= X + 6) && (yy >= Y && yy <= Y + 6))
            {
                activePoint = 0;
            }
            return activePoint != -1;
        }
        public void Draw(Graphics gr)
        {
            Pen pen = new Pen(Col);
            SolidBrush brush = new SolidBrush(Bcol);
            gr.FillEllipse(brush, (float)X, (float)Y, W, H);
            gr.DrawEllipse(pen, (float)X, (float)Y, W, H);
        }
    }
}
