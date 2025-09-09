// Util/TextUtil.cs
using System.Text;

namespace SingleTrainTrackScheduler.Util
{
    public static class TextUtil
    {
        // Türkçe karakterler için normalizasyon – basitçe NFC'ye getiriyoruz.
        public static string FixTR(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            return s.Normalize(NormalizationForm.FormC);
        }
    }
}
