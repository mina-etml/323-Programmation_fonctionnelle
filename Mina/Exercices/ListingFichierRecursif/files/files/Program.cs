using System;
using System.IO;

class Program
{
    static void Main()
    {
        Console.Write("Entrez le chemin du dossier à explorer : ");
        string chemin = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(chemin))
        {
            Console.WriteLine("Chemin invalide.");
            return;
        }

        if (!Directory.Exists(chemin) && !File.Exists(chemin))
        {
            Console.WriteLine("Le chemin spécifié n'existe pas.");
            return;
        }

        Console.WriteLine("\n=== Liste récursive des fichiers ===\n");
        ListerFichiers(chemin);
    }

    /// <summary>
    /// Liste récursivement les fichiers et dossiers d'un chemin donné.
    /// </summary>
    static void ListerFichiers(string chemin, int niveau = 0)
    {
        string indentation = new string(' ', niveau * 2);

        try
        {
            // Cas de base : c’est un fichier
            if (File.Exists(chemin))
            {
                Console.WriteLine($"{indentation}[F] {Path.GetFileName(chemin)}");
                return;
            }

            // Cas récursif : c’est un dossier
            if (Directory.Exists(chemin))
            {
                Console.WriteLine($"{indentation}[D] {Path.GetFileName(chemin)}");

                // Lister tous les fichiers du dossier
                foreach (string fichier in Directory.GetFiles(chemin))
                {
                    ListerFichiers(fichier, niveau + 1);
                }

                // Lister tous les sous-dossiers du dossier
                foreach (string sousDossier in Directory.GetDirectories(chemin))
                {
                    ListerFichiers(sousDossier, niveau + 1);
                }
            }
        }
        catch (UnauthorizedAccessException)
        {
            Console.WriteLine($"{indentation}[X] Accès refusé : {chemin}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"{indentation}[!] Erreur : {ex.Message}");
        }
    }
}
