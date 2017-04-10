namespace Dupire
{
    public interface IDupire
    {
        double LocalVolatility(double t, double logMoneyness);
    }
}