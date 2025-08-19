using System;
using System.Threading.Tasks;

namespace Kevsoft.Ssml
{
	public abstract class FluentSsml(ISsml inner) : ISsml
	{
		private readonly ISsml _inner = inner;

		IFluentSay ISsml.Say(string value) => _inner.Say(value);

		IFluentSayDate ISsml.Say(DateTime value) => _inner.Say(value);

		IFluentSayTime ISsml.Say(TimeSpan value) => _inner.Say(value);

		IFluentSayNumber ISsml.Say(int value) => _inner.Say(value);

		Task<string> ISsml.ToStringAsync() => _inner.ToStringAsync();

		IBreak ISsml.Break() => _inner.Break();

		ISsml ISsml.WithConfiguration(SsmlConfiguration configuration) => _inner.WithConfiguration(configuration);
	}
}
