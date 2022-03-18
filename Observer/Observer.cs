namespace Observer;

public enum Genre
{
    Sport,
    Politics,
    Economy,
    Science
}

public interface IObserver
{
    void Update(ISubject subject);
}

public interface ISubject
{
    void Attach(IObserver observer);
    void Detach(IObserver observer);
    void Notify();
}

public class NewsAgency : ISubject
{
    public string NewsHeadline;
    public Genre State;

    public void SetNewsHeadline(Genre state, string news)
    {
        this.NewsHeadline = news;
        this.State = state;
        Notify();
    }

    private List<IObserver> Observers = new List<IObserver>();

    public void Attach(IObserver observer)
    {
        this.Observers.Add(observer);
    }

    public void Detach(IObserver observer)
    {
        this.Observers.Remove(observer);
    }

    public void Notify()
    {
        foreach (var observer in Observers)
        {
            observer.Update(this);
        }
    }
}

class DailyEconomy : IObserver
{
    public void Update(ISubject subject)
    {
        if ((subject as NewsAgency).State == Genre.Economy)
        {
            Console.WriteLine($"DailyEconomy publikuje artykuł {(subject as NewsAgency).NewsHeadline}");
        }
    }
}

class NewYorkTimes : IObserver
{
    public void Update(ISubject subject)
    {
        Console.WriteLine($"NewYorkTimes publikuje artykuł {(subject as NewsAgency).NewsHeadline}");
    }
}

class NationalGeographic : IObserver
{
    public void Update(ISubject subject)
    {
        if ((subject as NewsAgency).State == Genre.Science)
        {
            Console.WriteLine($"NationalGeographic publikuje artykuł {(subject as NewsAgency).NewsHeadline}");
        }
    }
}