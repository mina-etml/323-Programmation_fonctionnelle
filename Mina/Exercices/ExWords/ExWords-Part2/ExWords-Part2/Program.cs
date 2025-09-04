using System;
using System.Collections.Generic;

class Program
{
    // Dictionnaire des fréquences des lettres en français
    static Dictionary<char, double> epsilonValue = new()
    {
        ['A'] = 0.07636,
        ['B'] = 0.00901,
        ['C'] = 0.03260,
        ['D'] = 0.03669,
        ['E'] = 0.14715,
        ['F'] = 0.01066,
        ['G'] = 0.00866,
        ['H'] = 0.00737,
        ['I'] = 0.07529,
        ['J'] = 0.00613,
        ['K'] = 0.00049,
        ['L'] = 0.05456,
        ['M'] = 0.02968,
        ['N'] = 0.07095,
        ['O'] = 0.05302,
        ['P'] = 0.03063,
        ['Q'] = 0.01362,
        ['R'] = 0.06553,
        ['S'] = 0.07948,
        ['T'] = 0.07244,
        ['U'] = 0.06311,
        ['V'] = 0.01838,
        ['W'] = 0.00074,
        ['X'] = 0.00427,
        ['Y'] = 0.00128,
        ['Z'] = 0.00326
    };

    // Fonction pour calculer la valeur Epsilon d’un mot
    static double Epsilon(string word, Dictionary<char, double> frequencies)
    {
        double value = 0;
        var occurencies = new Dictionary<char, int>();

        foreach (char c in word.ToUpper())
        {
            if (frequencies.ContainsKey(c))
            {
                if (occurencies.ContainsKey(c))
                    occurencies[c]++;
                else
                    occurencies[c] = 1;
            }
        }

        foreach (char c in word.ToUpper())
        {
            value += frequencies[c] / occurencies[c];
        }

        return value;
    }

    // Point d’entrée du programme
    static void Main()
    {
        string[] words = {
            "bonjour", "ABA", "zoo", "rare", "alpha",
            "test", "unique", "HELLO", "ZEBRA", "FANTASY",
            "ECOLE", "MATH", "SCIENCE", "VIE", "AMOUR"
        };

        Console.WriteLine("Mots avec Epsilon entre 0.5 et 0.95 :\n");

        foreach (var word in words)
        {
            double eps = Epsilon(word, epsilonValue);
            {
                Console.WriteLine($"{word} → Epsilon = {eps:F4}");
            }
        }
    }
}
