void Dictionary()
{
    List<string> frenchWords = new List<string>() { "Merci", "Hotdog", "Oui", "Non", "Désolé", "Réunion", "Manger", "Boire", "Téléphone", "Ordinateur", "Internet", "Email", "Sandwich", "Hello", "Taxi", "Hotel", "Gare", "Train", "Bus", "Métro", "Tramway", "Vélo", "Voiture", "Piéton", "Feu rouge", "Cédez", "Ralentir", "gauche", "droite", "Continuer", "Sandwich", "Retourner", "Arrêter", "Stationnement", "Parking", "Interdit", "Péage", "Trafic", "Route", "Rond-point", "Football", "Carrefour", "Feu", "Panneau", "Vitesse", "Tramway", "Aéroport", "Héliport", "Port", "Ferry", "Bateau", "Canot", "Kayak", "Paddle", "Surf", "Plage", "Mer", "Océan", "Rivière", "Lac", "Étang", "Marais", "Forêt", "Hello", "Montagne", "Vallée", "Plaine", "Désert", "Jungle", "Savane", "Volleyball", "Tundra", "Glacier", "Neige", "Pluie", "Soleil", "Nuage", "Vent", "Tempête", "Ouragan", "Tornade", "Séisme", "Tsunami", "Volcan", "Éruption", "Ciel" };

    var source = "words.txt";
    const int maxParent = 5;
    int parent = 0;
    while (!File.Exists(source) && parent < maxParent)
    {
        //handles start from ide
        source = $"../{source}";
        parent++;
    }

    var englishWordsLC = File.ReadAllLines(source).Select(word => word.ToLower());

    var same = frenchWords.Where(word => englishWordsLC.Contains(word.ToLower()));
    Console.WriteLine(String.Join(", ", same));
}

Dictionary();