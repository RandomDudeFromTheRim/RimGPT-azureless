namespace Kevsoft.Ssml
{
	public static class SsmlConfigurationExtensions
	{
		public static ISsml ForAlexa(this ISsml ssml) => ssml.WithConfiguration(new SsmlConfiguration(true));
	}
}
