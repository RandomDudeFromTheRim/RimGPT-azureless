using Kevsoft.Ssml;
using Newtonsoft.Json;
using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using UnityEngine;
using UnityEngine.Networking;
using Verse;
using Verse.Sound;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;

namespace RimGPT
{
	public struct TTSResponse
	{
		public int Error;
		public string Speaker;
		public int Cached;
		public string Text;
		public string tasktype;
		public string URL;
		public string MP3;
	}

	public class Voice
	{
		public string Name;
		public string DisplayName;
		public string LocalName;
		public string ShortName;
		public string Gender;
		public string Locale;
		public string LocaleName;
		public string[] StyleList;
		public string SampleRateHertz;
		public string VoiceType;
		public string Status;
		public string WordsPerMinute;

		public static Voice From(string shortName)
		{
			if (shortName == null || shortName == "")
				return null;
			return TTS.voices?.FirstOrDefault(v => v.ShortName == shortName);
		}
	}

	public class CoquiVoice
	{
		public string VoiceId;
		public string Name;
		public string Language;
		public string Gender;
		public string ModelUrl;
		
		public static CoquiVoice[] AvailableVoices = new[]
		{
			new CoquiVoice { VoiceId = "en_UK", Name = "English (UK)", Language = "en", Gender = "Female" },
			new CoquiVoice { VoiceId = "en_US", Name = "English (US)", Language = "en", Gender = "Female" },
			new CoquiVoice { VoiceId = "ru_RU", Name = "Russian", Language = "ru", Gender = "Female" },
			new CoquiVoice { VoiceId = "fr_FR", Name = "French", Language = "fr", Gender = "Female" },
			new CoquiVoice { VoiceId = "de_DE", Name = "German", Language = "de", Gender = "Female" }
		};
		
		public static CoquiVoice From(string voiceId) => AvailableVoices.FirstOrDefault(v => v.VoiceId == voiceId);
	}

	public class VoiceStyle
	{
		private VoiceStyle(string name, string value, string tooltip) { Name = name; Value = value; Tooltip = tooltip; }

		public string Name { get; private set; }
		public string Value { get; private set; }
		public string Tooltip { get; private set; }

		public static readonly VoiceStyle[] Values =
		[
			new VoiceStyle("Default", "default", null),
			new VoiceStyle("Advertisement Upbeat", "advertisement_upbeat", "Expresses an excited and high-energy tone for promoting a product or service"),
			new VoiceStyle("Affectionate", "affectionate", "Expresses a warm and affectionate tone, with higher pitch and vocal energy"),
			new VoiceStyle("Angry", "angry", "Expresses an angry and annoyed tone"),
			new VoiceStyle("Assistant", "assistant", "Expresses a warm and relaxed tone for digital assistants"),
			new VoiceStyle("Calm", "calm", "Expresses a cool, collected, and composed attitude when speaking"),
			new VoiceStyle("Chat", "chat", "Expresses a casual and relaxed tone"),
			new VoiceStyle("Cheerful", "cheerful", "Expresses a positive and happy tone"),
			new VoiceStyle("Customer Service", "customerservice", "Expresses a friendly and helpful tone for customer support"),
			new VoiceStyle("Depressed", "depressed", "Expresses a melancholic and despondent tone with lower pitch and energy"),
			new VoiceStyle("Disgruntled", "disgruntled", "Expresses a disdainful and complaining tone"),
			new VoiceStyle("Excited", "excited", "Expresses an upbeat and hopeful tone"),
			new VoiceStyle("Fearful", "fearful", "Expresses a scared and nervous tone"),
			new VoiceStyle("Friendly", "friendly", "Expresses a pleasant, inviting, and warm tone"),
			new VoiceStyle("Gentle", "gentle", "Expresses a mild, polite, and pleasant tone"),
			new VoiceStyle("Serious", "serious", "Expresses a strict and commanding tone"),
			new VoiceStyle("Sad", "sad", "Expresses a sorrowful tone"),
			new VoiceStyle("Whispering", "whispering", "Speaks very softly and make a quiet and gentle sound")
		];

		public static VoiceStyle From(string shortName) => Values.FirstOrDefault(s => s.Value == shortName);
	}

	public class TTS
	{
		public static string APIURL => $"https://{RimGPTMod.Settings.azureSpeechRegion}.tts.speech.microsoft.com/cognitiveservices";
		
		// Coqui TTS configuration
		private static string CoquiServerUrl => RimGPTMod.Settings.coquiServerUrl ?? "http://localhost:5002";
		private static Process coquiServerProcess = null;
		
		public static Voice[] voices = [];
		
		private static AudioSource audioSource = null;
		private static readonly object audioSourceLock = new();

		public static void StartCoquiServer()
		{
			try
			{
				if (coquiServerProcess != null && !coquiServerProcess.HasExited)
					return;
					
				string pythonPath = FindPython();
				if (pythonPath == null)
				{
					Logger.Error("Python not found. Coqui TTS requires Python.");
					return;
				}
				
				var startInfo = new ProcessStartInfo
				{
					FileName = pythonPath,
					Arguments = "-m TTS.server.server --port 5002 --model_name tts_models/en/ljspeech/tacotron2-DDC",
					UseShellExecute = false,
					CreateNoWindow = true,
					RedirectStandardOutput = true,
					RedirectStandardError = true
				};
				
				coquiServerProcess = Process.Start(startInfo);
				Logger.Message("Coqui TTS server started on port 5002");
				
				// Wait a bit for server to start
				System.Threading.Thread.Sleep(5000);
			}
			catch (Exception e)
			{
				Logger.Error($"Failed to start Coqui server: {e.Message}");
			}
		}
		
		private static string FindPython()
		{
			var possiblePaths = new[]
			{
				"python",
				"python3",
				"py",
				"/usr/bin/python3",
				"/usr/local/bin/python3",
				"C:\\Python39\\python.exe"
			};
			
			foreach (var path in possiblePaths)
			{
				try
				{
					var process = Process.Start(new ProcessStartInfo
					{
						FileName = path,
						Arguments = "--version",
						UseShellExecute = false,
						CreateNoWindow = true,
						RedirectStandardOutput = true
					});
					
					process.WaitForExit(1000);
					if (process.ExitCode == 0)
						return path;
				}
				catch { }
			}
			
			return null;
		}

		public static AudioSource GetAudioSource()
		{
			lock (audioSourceLock)
			{
				if (audioSource == null)
				{
					var gameObject = new GameObject("HarmonyOneShotSourcesWorldContainer");
					UnityEngine.Object.DontDestroyOnLoad(gameObject);
					gameObject.transform.position = Vector3.zero;
					var gameObject2 = new GameObject("HarmonyOneShotSource");
					gameObject2.transform.parent = gameObject.transform;
					gameObject2.transform.localPosition = Vector3.zero;
					audioSource = AudioSourceMaker.NewAudioSourceOn(gameObject2);
					audioSource.spatialBlend = 0f;
					audioSource.rolloffMode = AudioRolloffMode.Linear;
					audioSource.minDistance = 100000;
					audioSource.bypassEffects = true;
					audioSource.bypassListenerEffects = true;
					audioSource.bypassReverbZones = true;
					audioSource.ignoreListenerPause = true;
					audioSource.ignoreListenerVolume = true;
					audioSource.volume = 1;
				}
				return audioSource;
			}
		}

		public static async Task<T> DispatchFormPost<T>(string path, WWWForm form, bool addSubscriptionKey, Action<string> errorCallback)
		{
			using var request = form != null ? UnityWebRequest.Post(path, form) : UnityWebRequest.Get(path);
			if (addSubscriptionKey)
				request.SetRequestHeader("Ocp-Apim-Subscription-Key", RimGPTMod.Settings.azureSpeechKey);
			try
			{
				var asyncOperation = request.SendWebRequest();
				while (!asyncOperation.isDone && RimGPTMod.Running)
					await Task.Delay(200);
			}
			catch (Exception exception)
			{
				var error = $"Error communicating with Azure: {exception}";
				errorCallback?.Invoke(error);
				return default;
			}
			var response = request.downloadHandler.text;
			var code = request.responseCode;
			if (code >= 300)
			{
				var error = $"Got {code} response from OpenAI: {response}";
				errorCallback?.Invoke(error);
				return default;
			}
			try
			{
				return JsonConvert.DeserializeObject<T>(response);
			}
			catch (Exception)
			{
				Logger.Error($"Azure malformed output: {response}");
			}
			return default;
		}

		public static async Task<AudioClip> AudioClipFromCoqui(Persona persona, string text, Action<string> errorCallback)
		{
			try
			{
				// Coqui TTS API call
				var coquiUrl = $"{CoquiServerUrl}/api/tts";
				var jsonPayload = JsonConvert.SerializeObject(new
				{
					text = text,
					voice = persona.coquiVoice ?? "en_US",
					speed = persona.speechRate == "default" ? 1.0f : GetSpeedFromRate(persona.speechRate),
					pitch = persona.speechPitch == "default" ? 1.0f : GetPitchFromValue(persona.speechPitch),
					style = persona.coquiVoiceStyle ?? "neutral"
				});
				
				var bytes = Encoding.UTF8.GetBytes(jsonPayload);
				using var request = new UnityWebRequest(coquiUrl, "POST");
				request.uploadHandler = new UploadHandlerRaw(bytes);
				request.downloadHandler = new DownloadHandlerAudioClip(coquiUrl, AudioType.WAV);
				request.SetRequestHeader("Content-Type", "application/json");
				
				var asyncOperation = request.SendWebRequest();
				while (!asyncOperation.isDone && RimGPTMod.Running)
					await Task.Delay(100);
					
				if (request.result == UnityWebRequest.Result.Success)
				{
					var audioClip = DownloadHandlerAudioClip.GetContent(request);
					return audioClip;
				}
				else
				{
					errorCallback?.Invoke($"Coqui TTS Error: {request.error}");
					return null;
				}
			}
			catch (Exception e)
			{
				errorCallback?.Invoke($"Coqui TTS Exception: {e.Message}");
				return null;
			}
		}
		
		private static float GetSpeedFromRate(string rate)
		{
			return rate switch
			{
				"x-slow" => 0.8f,
				"slow" => 0.9f,
				"medium" => 1.0f,
				"fast" => 1.2f,
				"x-fast" => 1.4f,
				_ => 1.0f
			};
		}
		
		private static float GetPitchFromValue(string pitch)
		{
			return pitch switch
			{
				"x-low" => 0.8f,
				"low" => 0.9f,
				"medium" => 1.0f,
				"high" => 1.2f,
				"x-high" => 1.4f,
				_ => 1.0f
			};
		}

		public static async Task<AudioClip> AudioClipFromAzure(Persona persona, string path, string text, Action<string> errorCallback)
		{
			// Fall back to Coqui if Azure is configured but fails
			if (RimGPTMod.Settings.fallbackToCoqui)
			{
				return await AudioClipFromCoqui(persona, text, errorCallback);
			}
			
			var voice = persona.azureVoice;
			var style = persona.azureVoiceStyle;
			var styledegree = persona.azureVoiceStyleDegree;
			var rate = persona.speechRate;
			var pitch = persona.speechPitch;
			var xml = await new Ssml().Say(text).WithProsody(rate, pitch).AsVoice(voice, style, styledegree).ToStringAsync();
			if (Tools.DEBUG)
				Logger.Warning($"[{voice}] [{style}] [{styledegree}] [{rate}] [{pitch}] => {xml}");
			using var request = UnityWebRequest.Put(path, Encoding.Default.GetBytes(xml));
			using var downloadHandlerAudioClip = new DownloadHandlerAudioClip(path, AudioType.MPEG);
			request.method = "POST";
			request.SetRequestHeader("Ocp-Apim-Subscription-Key", RimGPTMod.Settings.azureSpeechKey);
			request.SetRequestHeader("Content-Type", "application/ssml+xml");
			request.SetRequestHeader("X-Microsoft-OutputFormat", "audio-16khz-64kbitrate-mono-mp3");
			request.downloadHandler = downloadHandlerAudioClip;
			try
			{
				var asyncOperation = request.SendWebRequest();
				while (!asyncOperation.isDone && RimGPTMod.Running)
					await Task.Delay(200);
				RimGPTMod.Settings.charactersSentAzure += text.Length;
			}
			catch (Exception exception)
			{
				var error = $"Error communicating with Azure: {exception}";
				errorCallback?.Invoke(error);
				return await AudioClipFromCoqui(persona, text, errorCallback);
			}
			var code = request.responseCode;
			if (Tools.DEBUG)
				Logger.Warning($"Azure => {code} {request.error}");
			if (code >= 300)
			{
				var error = $"Got {code} response from Azure: {request.error}";
				errorCallback?.Invoke(error);
				return await AudioClipFromCoqui(persona, text, errorCallback);
			}
			return await Main.Perform(() =>
			{
				var audioClip = downloadHandlerAudioClip.audioClip;
				return audioClip;
			});
		}

		public static async Task<AudioClip> DownloadAudioClip(string path, Action<string> errorCallback)
		{
			using var request = UnityWebRequestMultimedia.GetAudioClip(path, AudioType.MPEG);
			try
			{
				var asyncOperation = request.SendWebRequest();
				while (!asyncOperation.isDone && RimGPTMod.Running)
					await Task.Delay(200);
			}
			catch (Exception exception)
			{
				var error = $"Error communicating with Azure: {exception}";
				errorCallback?.Invoke(error);
				return default;
			}
			var response = request.downloadHandler.text;
			var code = request.responseCode;
			if (code >= 300)
			{
				var error = $"Got {code} response from Azure: {response}";
				errorCallback?.Invoke(error);
				return default;
			}
			return DownloadHandlerAudioClip.GetContent(request);
		}

		public static void TestKey(Persona persona, Action callback)
		{
			Tools.SafeAsync(async () =>
			{
				var text = "This is a test message";
				string error = null;
				
				// Start Coqui server automatically
				if (RimGPTMod.Settings.useCoquiTTS || RimGPTMod.Settings.fallbackToCoqui)
				{
					StartCoquiServer();
				}
				
				if (RimGPTMod.Settings.IsConfigured)
				{
					var prompt = "Say something random.";
					if (persona.personalityLanguage != "-")
						prompt += $" Your response must be in {persona.personalityLanguage}.";
					var dummyAI = new AI();
					var result = await dummyAI.SimplePrompt(prompt);
					text = result.Item1;
					error = result.Item2;
				}
				if (text != null)
				{
					AudioClip audioClip = null;
					
					if (RimGPTMod.Settings.useCoquiTTS)
					{
						// Use Coqui TTS directly
						audioClip = await AudioClipFromCoqui(persona, text, e => error = e);
					}
					else
					{
						// Try Azure first, fall back to Coqui if configured
						audioClip = await AudioClipFromAzure(persona, $"{APIURL}/v1", text, e => error = e);
					}
					
					if (audioClip != null)
					{
						var source = GetAudioSource();
						source.Stop();
						source.clip = audioClip;
						source.volume = RimGPTMod.Settings.speechVolume;
						source.Play();
					}
				}
				if (error != null)
					LongEventHandler.ExecuteWhenFinished(() =>
					{
						var dialog = new Dialog_MessageBox(error, null, null, null, null, null, false, callback, callback);
						Find.WindowStack.Add(dialog);
					});
				else
					callback?.Invoke();
			});
		}
		
		public static void StopCoquiServer()
		{
			try
			{
				if (coquiServerProcess != null && !coquiServerProcess.HasExited)
				{
					coquiServerProcess.Kill();
					coquiServerProcess = null;
					Logger.Message("Coqui TTS server stopped");
				}
			}
			catch (Exception e)
			{
				Logger.Error($"Error stopping Coqui server: {e.Message}");
			}
		}
	}
}
