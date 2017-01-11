using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OOSCommon
{
    public class selfAttribute : Attribute
    {
        public selfAttribute(string EN, string CN)
        {
            m_ENDisplayText = EN;
            m_CNDisplayText = CN;
        }

        private string m_ENDisplayText = string.Empty;
        private string m_CNDisplayText = string.Empty;
        public string ENDisplayText
        {
            get { return m_ENDisplayText; }
        }

        public string CNDisplayText
        {
            get { return m_CNDisplayText; }
        }
    }
}
