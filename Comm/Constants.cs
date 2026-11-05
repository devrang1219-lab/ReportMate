using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoborSmartReport.Comm
{
    public static class Constants
    {
        static public class Ini
        {
            // [Info]
            // StartTime=2026-05-13 13:35:16
            // FileName=.\ini\main.ini
            // [Config]

            public const string INI_FILE_NAME = "config.ini";
            public const string INI_SECTION_NAME = "Settings";
            public const string INI_KEY_NAME = "Value";

        }

        public class Gap
        {
            public float Left { get; set; }
            public float Top { get; set; }
            public float Right { get; set; }
            public float Bottom { get; set; }
        }

        public static readonly Gap GapMachine = new Gap
        {
            Left = 0.8f,
            Top = 1.3f,
            Right = 0f,
            Bottom = 0.4f
        };

        public static readonly Gap GapFever = new Gap
        {
            Left = 0.8f,
            Top = 1.3f,
            Right = 0f,
            Bottom = 0.4f
        };

        public static readonly List<string> HIGH_WAVE_GRAPH = new List<String>()
        {
            "고_그.pdf",
            "고그.pdf"
        };

        public static readonly List<string> TIME_WAVE_GRAPH = new List<String>()
        {
            "시_그.pdf",
            "시그.pdf"
        };

        public static readonly List<string> TEST_REPORT = new List<String>()
        {
            "시_보.pdf",
            "시보.pdf"
        };
    }
}
