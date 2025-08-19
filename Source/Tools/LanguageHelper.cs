using System;

namespace RimGPT;

public class LanguageHelper
{
	/// <summary>
	/// Calculates the Levenshtein distance between two strings, which represents the number of single-character edits
	/// (insertions, deletions, or substitutions) required to change one word into the other. This is often used to
	/// increase or decrease the frequencyPenalty for ChatGPT to hopefully avoid repetitiveness.
	/// </summary>
	/// <param name="source">The source string to compare.</param>
	/// <param name="target">The target string to compare.</param>
	/// <returns>An integer representing the Levenshtein distance between the two provided strings.</returns>

	public static int CalculateLevenshteinDistance(string source, string target)
	{
		if (string.IsNullOrEmpty(source))
		{
			return string.IsNullOrEmpty(target) ? 0 : target.Length;
		}

		if (string.IsNullOrEmpty(target))
		{
			return source.Length;
		}

		var lengthSource = source.Length;
		var lengthTarget = target.Length;
		var matrix = new int[lengthSource + 1, lengthTarget + 1];

		// Initialize the matrix.
		for (var i = 0; i <= lengthSource; matrix[i, 0] = i++) { }
		for (var j = 0; j <= lengthTarget; matrix[0, j] = j++) { }

		for (var i = 1; i <= lengthSource; i++)
		{
			for (var j = 1; j <= lengthTarget; j++)
			{
				var cost = (target[j - 1] == source[i - 1]) ? 0 : 1;
				matrix[i, j] = Math.Min(
					Math.Min(matrix[i - 1, j] + 1, matrix[i, j - 1] + 1),
					matrix[i - 1, j - 1] + cost
				);
			}
		}

		return matrix[lengthSource, lengthTarget];
	}
}
