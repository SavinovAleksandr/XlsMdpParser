using System.Collections.Generic;

namespace XlsxMdpParser;

public class MdpBuilder
{
	public string ShemeName { get; set; }

	public string ShemeNum { get; set; }

	public List<TNV> TnvList { get; set; }
}
