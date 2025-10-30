// See https://aka.ms/new-console-template for more information
using Microsoft.VisualBasic;
using System.Diagnostics;

class Program
{

    static int Fibonacci(int n)
    {
        if (n < 0) return 0; // Handles negative inputs

        if (n == 0) return 0;
        if (n == 1) return 1;

        return Fibonacci(n -1) + Fibonacci(n -2);
    }

     /*
    for (int i = 0; i < 13; i ++)
    {
        Console.WriteLine(Fibonacci(i));
    }
    */

    static int FibonacciOptimise(int iteration, int precedent = 0, int actuel = 1)
    {
        if (iteration == 0) return precedent;
        if (iteration == 1) return actuel;

        return FibonacciOptimise(iteration - 1, actuel, precedent + actuel);
    }

     /*
    for (int i = 0; i < 13; i++)
    {
        Console.WriteLine(FibonacciOptimise(i));
    }
    */

    static void Main()
    {
        var sw = new Stopwatch();

        Console.WriteLine("=== Test de performance Fibonacci ===\n");

        // --- Test 1 : Version naïve ---
        sw.Start();
        var resultNaive = Fibonacci(35);
        sw.Stop();
        Console.WriteLine($"Fibonacci(35) [Naïve] = {resultNaive}  | Temps = {sw.ElapsedMilliseconds} ms");

        // --- Test 2 : Version optimisée ---
        sw.Restart();
        var resultOpt = FibonacciOptimise(35);
        sw.Stop();
        Console.WriteLine($"Fibonacci(35) [Optimisée] = {resultOpt}  | Temps = {sw.ElapsedTicks} ticks ({sw.ElapsedMilliseconds} ms)");

        // --- Test 4 : Grand test FibonacciOptimise(10000) ---
        sw.Restart();
        var resultBig = FibonacciOptimise(10000);
        sw.Stop();
        Console.WriteLine($"\nFibonacci(10000) [Optimisée]  = {resultOpt} | Temps = {sw.ElapsedMilliseconds} ms");

        Console.WriteLine("\nTests terminés !");
    }
}