using cmgenerator.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CMGenerator.Helper
{
    public class AreaHelper
    {
        private static List<Area> _list = new List<Area>();

        public static Area GetArea(string name)
        {
            if (string.IsNullOrEmpty(name)) return null;

            var area = _list.Find(x => x.Name.ToUpper().Equals(name.Trim().ToUpper()));
            if (area != null) return area;

            area = new Area { Name = name };
            _list.Add(area);
            return area;
        }

    }
}
