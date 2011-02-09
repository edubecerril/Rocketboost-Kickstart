using System;
using System.Threading;
using System.ComponentModel;

namespace Demo
{
    public class Fibonacci
    {
        /// <summary>
        /// Generates the first 'n' numbers in the Fibonacci
        /// sequence.
        /// </summary>
        /// <param name="n">Number of terms in the sequence
        /// to generate.</param>
        /// <returns>First n numbers in the Fibonacci
        /// sequence.</returns>
        private int NthTerm(int n)
        {
            int[] results;

            // Limit the length of the sequence
            // that can be requested.
            if (n >= 1 && n <= 50)
            {
                results = new int[n];

                int auxiliar = 0;
                int previous = 0;
                int current = 1;
                int i = 0;
                while (i < n)
                {
                    if (i == 0)
                    {
                        results[i] = 0;
                    }
                    else if (i == 1)
                    {
                        results[i] = 1;
                    }
                    else
                    {
                        auxiliar = previous;
                        previous = current;
                        current = auxiliar + current;
                        results[i] = current;
                    }

                    i++;
                }
            }
            else
            {
                results = new int[1];

                results[0] = -1;
            }

            return results[n-1];
        }
    }
}
