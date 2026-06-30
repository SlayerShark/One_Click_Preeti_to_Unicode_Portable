using System;
using System.Text.RegularExpressions;

namespace WordPreetiToUnicode
{
    public static class PreetiEngine
    {
        public static string ConvertBlock(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            input = ResolveMultiGlyphs(input);
            input = MapToUnicodeWithPlaceholders(input);
            input = ReorderLeftVowelsSmart(input);

            return input;
        }

        private static string ResolveMultiGlyphs(string text)
        {
            text = text.Replace("cf]", "ओ");
            text = text.Replace("cf)", "औ");
            text = text.Replace("f]", "ो");
            text = text.Replace("f)", "ौ");
            return text;
        }

        private static string MapToUnicodeWithPlaceholders(string text)
        {
            string output = "";
            for (int i = 0; i < text.Length; i++)
            {
                // Isolate the left vowel modifier for programmatic relocation
                if (text[i] == 'l')
                {
                    output += "__IKAR__";
                    continue;
                }

                // Match two-character glyph structures
                if (i + 1 < text.Length)
                {
                    string doubleKey = text.Substring(i, 2);
                    if (PreetiMapping.Dictionary.ContainsKey(doubleKey))
                    {
                        output += PreetiMapping.Dictionary[doubleKey];
                        i++;
                        continue;
                    }
                }

                // Match single character glyph structures
                string singleKey = text[i].ToString();
                if (PreetiMapping.Dictionary.ContainsKey(singleKey))
                {
                    output += PreetiMapping.Dictionary[singleKey];
                }
                else
                {
                    output += singleKey;
                }
            }
            return output;
        }

        private static string ReorderLeftVowelsSmart(string text)
        {
            // Shifts the placeholder past simple or complex consonant clusters (\u0915-\u0939 covers क-ह)
            string pattern = @"__IKAR__((?:[\u0915-\u0939]्)*[\u0915-\u0939])";
            text = Regex.Replace(text, pattern, "$1ि");

            // Clean up standalone instances
            return text.Replace("__IKAR__", "ि");
        }
    }
}