namespace FP.Templating.Excel
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    static class QueueExtensions
    {
        public static IEnumerable<T> DequeueWhile<T>(this Queue<T> @this, Predicate<T> condition)
        {
            while(@this.Any() && condition(@this.Peek()))
            {
                yield return @this.Dequeue();
            }
        }
    }
}
