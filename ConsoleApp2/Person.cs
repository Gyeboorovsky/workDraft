namespace ConsoleApp2;

public class Person
{
    public string Name { get; set; }

    private Person[] _backingStore;

    public Person this[int index]
    {
        get
        {
            return _backingStore[index];
        }
        set
        {
            _backingStore[index] = value;
        }
    }
}