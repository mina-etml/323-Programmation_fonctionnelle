using System.Text.Json;
/*
var client = new HttpClient();
var json = await HttpGetAsync(client, "films");
//Console.WriteLine(json);

var moviesResult = JsonSerializer.Deserialize<FilmResult>(json);

var movies = moviesResult.results; //List<Film>
//Console.WriteLine(movies[0].title);

await Planete1();
*/

class Program
{
    static async Task Main(string[] args)
    {
        var client = new HttpClient();
        var json = await HttpGetAsync(client, "films");

        var moviesResult = JsonSerializer.Deserialize<FilmResult>(json);
        var movies = moviesResult.results;

        await Planete1(client, movies);
    }

    static async Task Planete1(HttpClient client, List<Film> movies)
    {
        //Titre le plus long
        var longestMovieTitle = movies.Select(film => film.title).Aggregate((title1, title2) => (title1.Length > title2.Length) ? title1 : title2);
        Console.WriteLine(longestMovieTitle);

        //Personnage présent dans le plus de film
        //Select gets a list of list of film characters, SelectMany flattens it to just a list of film characters
        var mostReccurentCharacterUrl = movies.SelectMany(film => film.characters).GroupBy(characterUrl => characterUrl).OrderByDescending(g => g.Count()).First().Key;
        var mostReccurentCharacterJson = await HttpGetAsync(client, mostReccurentCharacterUrl);
        var mostReccurentCharacter = JsonSerializer.Deserialize<Character>(mostReccurentCharacterJson);
        Console.WriteLine($"Nom personnage le plus fréquent: {mostReccurentCharacter.name}");

        //Planète la plus peuplée
        var mostPeupledPlanet = movies.SelectMany(film => film.planets).GroupBy(planetsUrl => planetsUrl).OrderByDescending(g => g.Count()).First().Key;
        var mostReccurentCharacterJson = await HttpGetAsync(client, mostReccurentCharacterUrl);
        var mostReccurentCharacter = JsonSerializer.Deserialize<Character>(mostReccurentCharacterJson);
        Console.WriteLine($"Nom personnage le plus fréquent: {mostReccurentCharacter.name}");


    }



    static async Task<string> HttpGetAsync(HttpClient client, string query)
    {
        var response = await client.GetAsync(query.Contains("https") ? query : "https://swapi.dev/api/" + query);
        response.EnsureSuccessStatusCode();
        var json = await response.Content.ReadAsStringAsync();

        return json;
    }

}

public static class Extension
{
    public static void Write(this IEnumerable<object> target, char separator=',')
    {
        Console.WriteLine(String.Join(separator, target));
    }
}

class FilmResult
{
    public int count { get; set; }
    public List<Film> results { get; set; }
}

class Film
{
    public string title { get; set; }
    public List<string> characters { get; set; }

}

public class Character
{
    public string name { get; set; }
}

public class Planets
{
    public double population { get; set; }
}
