using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ZubApproxProject
{
    public class HurstExponent
    {
        public HurstExponent()
        {
        }
        public double ComputeHurstExponent(double[] inputArr)
        {
            int count = inputArr.Length;
            //find average
            double averageValue = inputArr.Sum() / count;

            double[] partialSumms = new double[count];
            double summ = 0.0, summarySquaredValue = 0.0;
            for (int i = 0; i < count; ++i)
            {
                double value = inputArr[i] - averageValue;
                partialSumms[i] = (summ += value);
                summarySquaredValue += (value * value);
            }

            double R = partialSumms.Max() - partialSumms.Min(); //range
            double S = Math.Sqrt(summarySquaredValue / count); //standard deviation 

            return Math.Log(R / S) / Math.Log(count / 2);
        }
    }
}
