﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace APILib.APIOptions
{
    public class SplitterAPIOptions: APIOptionsBase
    {
        public Guid FileId { get; set; }

        /// <summary>
        /// Количество строк/файлов в новом созданном файле
        /// </summary>
        public int ResultsCount { get; set; }

        /// <summary>
        /// Количество строк, которые требуется использовать как заголовки 
        /// </summary>
        [DefaultValue(0)]
        public int AddHeaderRows { get; set; } = 0;

        /// <summary>
        /// Критерий, по которому происходит разделение данных 
        /// </summary>
        public enum SplitType
        {
            SplitByRows,
            SplitByFiles
        }

        [DefaultValue(SplitType.SplitByRows)]
        public SplitType SplitMode { get; set; } = SplitType.SplitByRows;
    }
}
