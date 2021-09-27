using System.Drawing;

namespace KursachExcelInterop
{
    public class ImageData
    {
        public ImageData(PointF point, int checkRes)
        {
            Point = point;
            CheckRes = checkRes;
            Triggerctangle = new Rectangle((int)point.X - 4, (int)point.Y - 4, 7, 7);
        }

        public PointF Point { get; }
        public int CheckRes { get; }
        public Rectangle Triggerctangle { get; }
    }
}