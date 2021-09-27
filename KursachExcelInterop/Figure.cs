using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace KursachExcelInterop
{
    class Figure
    {
        private double x, y;
        private int h, w;
        private Color col, backColor;

        public Figure()
        {
            x = y = h = w = 1;
            col = backColor = Color.Black;
        }

        public Figure(double x, double y, int h, int w, Color col, Color backColor)
        {
            this.x = x;
            this.y = y;
            this.h = h;
            this.w = w;
            Col = col;
            Bcol = backColor;
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
        public Color Bcol
        {
            get
            {
                return backColor;
            }
            set
            {
                backColor = value;
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
        public int W
        {
            get
            {
                return w;
            }
            set
            {
                if (value > 0)
                    w = value;
                else
                    w = 1;
            }
        }
        public int H
        {
            get
            {
                return h;
            }
            set
            {
                if (value > 0)
                    h = value;
                else
                    h = 1;
            }
        }
    }
}
