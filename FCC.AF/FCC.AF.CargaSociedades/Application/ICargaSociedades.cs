using Microsoft.Extensions.Logging;

namespace FCC.AF.CargaSociedades.Application
{
    public interface ICargaSociedades
    {
        void DoProcess(ILogger log);
    }
}

