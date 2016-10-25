using System;
using System.IO;
using System.Speech.Recognition;

namespace EllisMIS.Audio.Transcription.Microsoft
{
    public class Dictation : IDisposable
    {
        #region Local Variables
        private SpeechRecognitionEngine _speechRecognitionEngine = null;
        private DictationGrammar _dictationGrammar = null;
        private bool _disposed = false;
        #endregion

        #region Constructors

        public Dictation()
        {
            ConstructorSetup();
        }

        public Dictation(DictationGrammar targetGrammar)
        {
            _dictationGrammar = targetGrammar;
            ConstructorSetup();
        }

        #endregion

        /// <summary>
        /// Start the transcriber using your default microphone.
        /// </summary>
        public void Start()
        {

            _speechRecognitionEngine.SetInputToDefaultAudioDevice();
            StartSetup();
        }

        /// <summary>
        /// Transcribe a .wav file
        /// </summary>
        /// <param name="targetWavFile"></param>
        public void Start(string targetWavFile)
        {
            if (!File.Exists(targetWavFile))
            {
                throw new FileNotFoundException("Specified WAV file does not exist.", "targetWavFile");
            }

                _speechRecognitionEngine.SetInputToWaveFile(targetWavFile);


            StartSetup();
        }

        private void StartSetup()
        {
            if (_dictationGrammar == null)
            {
                _dictationGrammar = new DictationGrammar();
            }

            _speechRecognitionEngine.LoadGrammar(_dictationGrammar);
            _speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);

            _speechRecognitionEngine.SpeechRecognized -= new EventHandler<SpeechRecognizedEventArgs>(SpeechRecognized);
            _speechRecognitionEngine.SpeechHypothesized -= new EventHandler<SpeechHypothesizedEventArgs>(SpeechHypothesizing);

            _speechRecognitionEngine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(SpeechRecognized);
            _speechRecognitionEngine.SpeechHypothesized += new EventHandler<SpeechHypothesizedEventArgs>(SpeechHypothesizing);
        }

        public void Stop()
        {
            _speechRecognitionEngine.UnloadGrammar(_dictationGrammar);
            _speechRecognitionEngine.RecognizeAsyncStop();
        }

        private void ConstructorSetup()
        {
            _speechRecognitionEngine = new SpeechRecognitionEngine();

            _speechRecognitionEngine.UnloadAllGrammars();

       }

        private void SpeechHypothesizing(object sender, SpeechHypothesizedEventArgs e)
        {
            OnSpeechHypothesizingEvent(e);
        }

        private void SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            OnSpeechRecognizedEvent(e);
        }

        public delegate void SpeechRecognizedEventHandler(object sender, SpeechRecognizedEventArgs e);
        public delegate void SpeechHypothesizingEventHandler(object sender, SpeechHypothesizedEventArgs e);

        /// <summary>
        /// The final results of the audio transcription
        /// </summary>
        public event SpeechRecognizedEventHandler SpeechRecognizedEvent;

        /// <summary>
        /// The real time guessing of the audio transcription
        /// </summary>
        public event SpeechHypothesizingEventHandler SpeechHypothesizingEvent;

        protected void OnSpeechRecognizedEvent(SpeechRecognizedEventArgs e)
        {
            if (SpeechRecognizedEvent != null)
            {
                SpeechRecognizedEvent(this, e);
            }
        }
        protected void OnSpeechHypothesizingEvent(SpeechHypothesizedEventArgs e)
        {
            if (SpeechHypothesizingEvent != null)
            {
                SpeechHypothesizingEvent(this, e);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this._disposed)
            {
                // If disposing equals true, dispose all managed
                // and unmanaged resources.
                if (disposing)
                {
                    Stop();
                    _speechRecognitionEngine.SpeechRecognized -= new EventHandler<SpeechRecognizedEventArgs>(SpeechRecognized);
                    _speechRecognitionEngine.SpeechHypothesized -= new EventHandler<SpeechHypothesizedEventArgs>(SpeechHypothesizing);

                    _dictationGrammar = null;
                    _speechRecognitionEngine.UnloadAllGrammars();
                    _speechRecognitionEngine.Dispose();
                }

                // Note disposing has been done.
                _disposed = true;

            }
        }

    }
}
