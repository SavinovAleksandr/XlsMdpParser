using System.Collections.Generic;

namespace XlsxMdpParser;

public class TNV
{
	public string Tnv { get; set; }

	public List<MDP> MdpNoPA { get; set; }

	public List<MDP> MdpPa { get; set; }

	public string Adp { get; set; }

	public List<MDP> MdpNoPaCriteria { get; set; }

	public List<MDP> MdpPaCriteria { get; set; }

	public string AdpCriteria { get; set; }

	public List<string> MdpNoPaDop { get; set; }

	public List<string> MdpPaDop { get; set; }

	public List<string> AdpDop { get; set; }
}
