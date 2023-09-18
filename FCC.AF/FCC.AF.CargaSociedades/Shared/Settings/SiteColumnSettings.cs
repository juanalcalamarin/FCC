using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FCC.AF.CargaSociedades.Shared.Settings
{
    public class SiteColumnSettings
    {
        public string Title { get; set; }
        public string Display { get; set; }
        public string TipoDato { get; set; }
        public string Tipo { get; set; }
        public bool? Statutory { get; set; }
        public int? StatutoryExcelColumn { get; set; }
        public bool? Reporting { get; set; }
        public int? ReportingExcelColumn { get; set; }
        public bool? Certificates { get; set; }
        public int? CertificatesExcelColumn { get; set; }

        public SiteColumnSettings()
        { }

        public SiteColumnSettings(
            string _title,
            string _display,
            string _tipodato,
            string _tipo,
            bool _statutory,
            int _statutorycolumn,
            bool _reporting,
            int _reportingcolumn,
            bool _certificates,
            int _certificatescolumn)
        {
            Title = _title;
            Display = _display;
            TipoDato = _tipodato;
            Tipo = _tipo;
            Statutory = _statutory;
            StatutoryExcelColumn = _statutorycolumn;
            Reporting = _reporting;
            ReportingExcelColumn = _reportingcolumn;
            Certificates = _certificates;
            CertificatesExcelColumn = _certificatescolumn;
        }
    }
}
