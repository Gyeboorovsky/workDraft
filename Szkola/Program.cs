using System;
using System.Collections.Generic;

public abstract class ProductPrototype
{
    public decimal Price;
    public ProductPrototype(decimal price)
    {
        Price = price;
    }

    public ProductPrototype Clone()
    {
        return (ProductPrototype)this.MemberwiseClone();
    }
}

public class Bread : ProductPrototype
{
    public Bread(decimal price) : base(price) { }
}

public class Butter : ProductPrototype
{
    public Butter(decimal price) : base(price) { }
}

public class Supermarket
{
    private Dictionary<string, ProductPrototype> _productList = new Dictionary<string, ProductPrototype>();
    public void AddProduct(string key, ProductPrototype productPrototype)
    {
        _productList.Add(key, productPrototype);
    }

    public ProductPrototype GetClonedProduct(string key)
    {
        return _productList[key].Clone();
    }
}

class MainClass
{
    public static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Supermarket supermarket = new Supermarket();
        ProductPrototype product;
        supermarket.AddProduct("Bread", new Bread(9.50m));
        supermarket.AddProduct("Butter", new Butter(5.30m));
        product = supermarket.GetClonedProduct("Bread");
        decimal product2 = (product.Price * 0.9m);

        Console.WriteLine(String.Format("Bread - {0} zł", product.Price + " > " + product2.ToString("F2")));
        Console.WriteLine();

        product = supermarket.GetClonedProduct("Bread");
        Console.WriteLine(String.Format("Bread - {0} zł", product.Price));
        Console.WriteLine();

        product = supermarket.GetClonedProduct("Butter");
        Console.WriteLine(String.Format("Butter - {0} zł", product.Price));

    }
}