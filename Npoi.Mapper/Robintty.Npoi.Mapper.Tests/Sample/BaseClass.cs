using Robintty.Npoi.Mapper.Attributes;

namespace Robintty.Npoi.Mapper.Tests.Sample
{
    /// <summary>
    /// The base class for sample classes.
    /// </summary>
    public class BaseClass
    {
        public string BaseStringProperty { get; set; }

        [Ignore]
        public string BaseIgnoredProperty { get; set; }
    }
}
