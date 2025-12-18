using System;
using System.Collections.Generic;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading.Tasks;

namespace PricingSheet.Alerts
{
    public class SpeechAlerts
    {
        public int Volume { get; set; }
        public int Rate { get; set; }
        public VoiceGender Gender { get; set; }
        public VoiceAge Age { get; set; }

        public SpeechAlerts(int volume = 100, int rate = -2, VoiceGender gender = VoiceGender.Male, VoiceAge age = VoiceAge.Adult)
        {
            Volume = volume > 100 ? 100 : volume < 0 ? 0 : volume;
            Rate = rate > 10 ? 10 : rate < -10 ? -10 : rate;
            Gender = gender;
            Age = age;
        }

        public void Speak(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
                return;

            Task.Run(() =>
            {
                using (SpeechSynthesizer synthesizer = new SpeechSynthesizer())
                {
                    synthesizer.SelectVoiceByHints(Gender, Age);
                    synthesizer.Volume = Volume;
                    synthesizer.Rate = Rate;
                    synthesizer.Speak(message);
                }
            });
        }
    }
}
