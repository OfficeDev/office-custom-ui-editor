using System;
using System.Collections;
using System.Text;
using System.Text.RegularExpressions;

using System.Xml;

namespace CustomUIEditor
{
	/// <summary>
	/// Very basic, read: inefficient, Xml colorizer based on RegEx.
	/// </summary>
	internal class XmlColorizer
	{
		public static string Colorize(string text)
		{
            text = ConvertUTFToRtf(text);

			if (text == null || text.Length == 0)
			{
				return rtfString;
			}

			StringBuilder rtf = new StringBuilder(3 * text.Length);
			rtf.Append(rtfString);

			text = text.Replace("{", "\\{");
			text = text.Replace("}", "\\}");
			rtf.Append(ParseXmlComment(text));

            return rtf.ToString().Replace("\n", @"\par ").Replace("\\\\u", "\\u"
                        /*jargil: there is an issue with the ToString: it adds an extra \ to every \ in the StringBuilder object,
                         * that's why we need to erase it from the returning string.*/);
		}

		private static Regex xmlCommentRegex = new Regex(
			@"<!--(.*?)-->",
			RegexOptions.IgnoreCase
			| RegexOptions.Singleline
			| RegexOptions.CultureInvariant
			| RegexOptions.IgnorePatternWhitespace
			| RegexOptions.Compiled
			);

		private static string ParseXmlComment(string text)
		{
			StringBuilder rtfText = new StringBuilder(3 * text.Length);
			MatchCollection matches = xmlCommentRegex.Matches(text);
			if (matches.Count == 0)
			{
				rtfText.Append(ParseXmlAttribute(text));
			}
			else
			{
				int vCurrent = 0;
				foreach (Match match in matches)
				{
					rtfText.Append(ParseXmlAttribute(text.Substring(vCurrent, match.Index - vCurrent)));

					rtfText.Append(rtfDelimiter + "<!--" + rtfComment);
					rtfText.Append(match.Groups[1].Value);
					rtfText.Append(rtfDelimiter + "-->");
					vCurrent = match.Index + match.Length;
				}
				rtfText.Append(ParseXmlAttribute(text.Substring(vCurrent)));
			}
			return rtfText.ToString();
		}

		private static Regex attributePairRegex = new Regex(
			@"(?<Keyword>\w+)(?<EqualSign>\s*=\s*)""(?<Value>.*?)""",
			RegexOptions.IgnoreCase
			| RegexOptions.Multiline
			| RegexOptions.CultureInvariant
			| RegexOptions.IgnorePatternWhitespace
			| RegexOptions.Compiled
			);

		private static string ParseXmlAttribute(string text)
		{
			StringBuilder rtfText = new StringBuilder(3 * text.Length);
			MatchCollection matches = attributePairRegex.Matches(text);
			if (matches.Count == 0)
			{
				rtfText.Append(ParseXmlTag(text));
			}
			else
			{
				int vCurrent = 0;
				foreach (Match match in matches)
				{
					rtfText.Append(ParseXmlTag(text.Substring(vCurrent, match.Index - vCurrent)));

					rtfText.Append(rtfAttributeName);
					rtfText.Append(match.Groups["Keyword"].Value);
					rtfText.Append(rtfDelimiter);
					rtfText.Append(match.Groups["EqualSign"].Value);
					rtfText.Append(rtfAttributeQuote);
					rtfText.Append("\"");
					rtfText.Append(rtfAttributeValue);
					rtfText.Append(match.Groups["Value"].Value);
					rtfText.Append(rtfAttributeQuote);
					rtfText.Append("\"");

					rtfText.Append(rtfDelimiter);
					vCurrent = match.Index + match.Length;
				}
				rtfText.Append(ParseXmlTag(text.Substring(vCurrent)));
			}
			return rtfText.ToString();
		}

		private static string ParseXmlTag(string text)
		{
			StringBuilder rtfText = new StringBuilder(2 * text.Length);
			for (int i = 0; i < text.Length; i++)
			{
				switch (text[i])
				{
					case '>':
						rtfText.Append(rtfDelimiter + text[i]);
						break;
					case '/':
					case '<':
						rtfText.Append(rtfDelimiter + text[i] + rtfName);
						break;
					case '\\':
						rtfText.Append("\\\\"); // JArgil:  This solves a bug where if you type \ you loose a line
						break;
					default:
						rtfText.Append(text[i]);
						break;
				}
			}
			return rtfText.ToString();
		}

        /// <summary>
        /// Converts from UTF to Rtf.
        /// </summary>
        /// <param name="unicode">String with UTF characters.</param>
        /// <returns>String with Rtf formatting.</returns>
        private static string ConvertUTFToRtf(string unicode)
        {
            System.Text.StringBuilder rtf = new System.Text.StringBuilder();
            foreach (char letter in unicode)
            {
                if (letter <= 0x7F) //before this is ASCII in UTF-8 and UTF-16 Encoding
                {
                    rtf.Append(letter);
                }
                else //starts Eurpean (except ASCII), arabic, hebrew, etc.
                {
                    rtf.Append(string.Format(@"\u{0}?", Convert.ToUInt32(letter)));
                }
            }
            return rtf.ToString(/*it has the same text but the utf characters where changed to something like \\u###?*/);
        }

		internal const string rtfString = @"{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fmodern\fprq1\fcharset0 Courier New;}}{\colortbl	;\red0\green0\blue255;\red128\green0\blue0;\red255\green0\blue0;\red0\green128\blue0;}\pard\f0\fs20 ";
		internal const string rtfAttributeName = @"\cf3 ";
		internal const string rtfAttributeValue = @"\cf1 ";
		internal const string rtfDelimiter = @"\cf1 ";
		internal const string rtfAttributeQuote = @"\cf0 ";
		internal const string rtfName = @"\cf2 ";
		internal const string rtfComment = @"\cf4 ";
	}
}
