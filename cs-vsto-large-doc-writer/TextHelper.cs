using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace LargeDocWriter
{
    public class TextHelper
    {
        public static bool IsRegularText(int part_type)
        {
            return part_type == 0;
        }

        public static bool IsCrossReference(int part_type)
        {
            return part_type == 1;
        }

        public static List<string> SplitText(string text, string regex_pattern, out List<int> part_types, out Dictionary<string, string> match_info_map1, out Dictionary<string, string> match_info_map2)
        {
            List<string> parts = new List<string>();
            part_types = new List<int>();

            MatchCollection mc = Regex.Matches(text, regex_pattern);

            match_info_map1 = new Dictionary<string, string>();
            match_info_map2 = new Dictionary<string, string>();

            if (mc.Count > 1)
            {
                List<string> groups = new List<string>();
                
                foreach (Match m in mc)
                {

                    if (m.Groups.Count > 0)
                    {
                        groups.Add(m.Groups[0].Value);
                        if (m.Groups.Count > 1)
                        {
                            match_info_map1[m.Groups[0].Value] = m.Groups[1].Value;
                        }
                        if (m.Groups.Count > 2)
                        {
                            match_info_map2[m.Groups[0].Value] = m.Groups[2].Value;
                        }
                    }
                }

                for (int i = 0; i < groups.Count; ++i)
                {
                    string sub_text = text.Substring(0, text.IndexOf(groups[i]));
                    text = text.Substring(text.IndexOf(groups[i]) + groups[i].Length);
                    if (!string.IsNullOrEmpty(sub_text))
                    {
                        parts.Add(sub_text);
                        part_types.Add(0);
                    }
                    parts.Add(groups[i]);
                    part_types.Add(1);
                }
                if (!string.IsNullOrEmpty(text))
                {
                    parts.Add(text);
                    part_types.Add(0);
                }
            }

            return parts;
        }
        
    }
}
