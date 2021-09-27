using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KursachExcelInterop
{
    class DISKRF
    {
        private double x, y;
        private Color col;

        public DISKRF(double x, double y, Color col)
        {
            this.x = x;
            this.y = y;           
            Col = col;           
        }
        public Color Col
        {
            get
            {
                return col;
            }
            set
            {
                col = value;
            }
        }      
        public double Y
        {
            get
            {
                return y;
            }
            set
            {
                y = value;
            }
        }
        public double X
        {
            get
            {
                return x;
            }
            set
            {
                x = value;
            }
        }
        //public void Draw(Graphics gr)
        //{
        //    Pen pen = new Pen(Col);
        //    gr.DrawLine(pen,x1,y1,x2,y2);
        //    gr.FillEllipse(brush, (float)X, (float)Y, W, H);
        //    gr.DrawEllipse(pen, (float)X, (float)Y, W, H);
        //}
    }
}
