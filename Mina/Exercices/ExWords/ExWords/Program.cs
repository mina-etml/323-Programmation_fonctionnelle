// A. Filtrage basique

//Liste donnée
using System.Text.RegularExpressions;

string[] words = { "bonjour", "hello", "monde", "vert", "rouge", "bleu", "jaune" };


// Fonction de filtre
Func<string, bool> notContainsX = word => !word.Contains("x");
Func<string, bool> fourAtLeat = word => word.Length >= 4;
var average = words.Average(word => word.Length);
Func<string, bool> sameAsAvg = word => word.Length == average;


//Recueil de fonctions
var filters = new List<Func<string, bool>>{ notContainsX , fourAtLeat , sameAsAvg };

//Fonction pour l'ordre des éléments
Func<IEnumerable<string>, IEnumerable<string>> reversed = list => list.Reverse();
Func<IEnumerable<string>, IEnumerable<string>> alphabeticalOrder = list => list.OrderBy(element => element);
Func<IEnumerable<string>, IEnumerable<string>> nonAlphabeticalOrder = list => list.OrderByDescending(element => element);

//Receuil de l'ordre
var order = new List<Func<IEnumerable<string>, IEnumerable<string>>> { reversed, alphabeticalOrder, nonAlphabeticalOrder };



//Menu - choix du filtre
Console.WriteLine($"Liste de mots : {String.Join(',', words)}");
Console.WriteLine("1. Pas de x v1");
Console.WriteLine("2. >= 4");
Console.WriteLine("3. = moyenne de longueur dans la liste");
Console.Write("\nChoix: ");

int listChoice = Convert.ToInt32(Console.ReadLine()) - 1;


//Menu - choix de l'ordre
Console.WriteLine($"Ordonnée par ");
Console.WriteLine("1. Ordre inverse calculé");
Console.WriteLine("2. Triés de a-z");
Console.WriteLine("3. Triés de z-a");
Console.Write("\nChoix: ");


int orderChoice = Convert.ToInt32(Console.ReadLine()) - 1;

// Applique filter and order
var filtered = words.Where(filters[listChoice]);
var ordered = order[orderChoice](filtered);

//Affiche le résultat
Console.WriteLine($"\nRésultat: {String.Join(", ", ordered)}");

///////////////////////////////////////////////////////////////////////////////////////////////
// B. Données parasites 1

string[] wordsB = {
    "whatThe!!!", // parasite
    "bonjour", "hello", "monde", "vert", "rouge", "bleu", "jaune",
    "My kingdom for a horse !", "Ooops I did it again" // parasites
};

// Filtrage des parasites
var cleanedWordsB = wordsB.Skip(1).SkipLast(2);

// Affichage
Console.WriteLine($"B) Mots nettoyés : {String.Join(", ", cleanedWordsB)}\n");

///////////////////////////////////////////////////////////////////////////////////////////////
// C. Données parasites 2
string[] wordsC = { "+++++", "<<<<<", ">>>>>", "bonjour", "he&llo", "@@@@", "vert", "rouge", "bleu", "jaune", "#####", "%%%%%%%" };

// Filtrage des parasites
//var cleanedWordsC = wordsC.SkipWhile(element => !Regex.IsMatch(element, "^[a-zA-Z]")); //ceci ne s'occupe que des problèmes du début
var cleanedWordsC = wordsC.Where(element => Regex.IsMatch(element, @"^[a-zA-Z]+$"));


// Affichage
Console.WriteLine($"C) Mots nettoyés : {String.Join(", ", cleanedWordsC)}");


///////////////////////////////////////////////////////////////////////////////////////////////
// D. Elitisme

string[] wordsD = { "i am the winner", "hello", "monde", "vert", "rouge", "bleu", "i am the looser" };
var wordsList = wordsD.ToList(); 

int indexWinner = wordsList.FindIndex(word => word.Contains("winner"));
int indexLooser = wordsList.FindIndex(word => word.Contains("looser"));


Console.WriteLine($"The winner is: {wordsD[indexWinner]} ");
Console.WriteLine($"The loser is: {wordsD[indexLooser]}");