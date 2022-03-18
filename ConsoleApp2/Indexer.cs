namespace ConsoleApp2;

public class Indexer
{
    public Person Method(string name1, string name2)
    {
        var p = new Person();
        p[0] = new Person() {Name = name1};
        p[1] = new Person() {Name = name2};

        return p;
    }
}

