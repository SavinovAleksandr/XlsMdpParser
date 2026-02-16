using System.Collections.Generic;
using System.Text;

namespace XlsxMdpParser;

public class BracketNode
{
	public List<object> ContentParts { get; } = new List<object>();

	public override string ToString()
	{
		return ToStringInternal(0);
	}

	private string ToStringInternal(int indent)
	{
		StringBuilder stringBuilder = new StringBuilder();
		string text = new string(' ', indent * 2);
		stringBuilder.AppendLine(text + "Node:");
		foreach (object contentPart in ContentParts)
		{
			if (contentPart is string text2)
			{
				stringBuilder.AppendLine(text + "  Text: \"" + text2 + "\"");
			}
			else if (contentPart is BracketNode bracketNode)
			{
				stringBuilder.AppendLine(bracketNode.ToStringInternal(indent + 1));
			}
		}
		return stringBuilder.ToString();
	}
}
