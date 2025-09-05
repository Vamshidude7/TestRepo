public class Calculator
{
    public static int Multiply(int a, int b)
    {
        return a * b;
    }

    public int Add(int a, int b)
    {
        return a + b;
    }
    public static void main(string[] args)
    {
        Calculator calc = new Calculator();
        int sum = calc.Add(5, 10);
        int product = Multiply(5, 10);
        System.Console.WriteLine("Sum: " + sum);
        System.Console.WriteLine("Product: " + product);
    }
}