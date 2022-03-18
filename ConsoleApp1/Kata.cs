namespace ConsoleApp1;

public static class Kata
{
    public static List<string> Anagrams(string word, List<string> words)
    {
        var check = word;
        var anagrams = new List<string>();
        foreach (var w in words)
        {
            if (w.Length != check.Length)
            {
                break;
            }
            
            var current = w.ToCharArray();
            var c = check.ToCharArray().ToList();
            for (int i = 0; i < w.Length; i++)
            {
                if (c.Contains(current[i]))
                    c.Remove(current[i]);
            }
            if (c.Count == 0)
                anagrams.Add(w);

            check = word;
        }

        foreach (var a in anagrams)
        {
            Console.WriteLine(a);
        }
        return anagrams;
    }
}