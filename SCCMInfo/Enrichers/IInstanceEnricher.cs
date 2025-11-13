using Spectre.Console;

using System.Management;
using System.Text;

namespace SCCMInfo.Enrichers
{
    internal interface IInstanceEnricher
    {
        string ClassName { get; }

        void Enrich(
            ManagementBaseObject targetInstance,
            Table table,
            StringBuilder logMessage,
            ManagementScope scope);
    }
}
