using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExportToExcel.Helpers
{
    public static class ExtendedLinq
    {
        public static void ForEach<T>(this IEnumerable<T> source, Action<T> action) =>
            source.ForEach(action, CancellationToken.None);

        public static void ForEach<T>(this IEnumerable<T> source, Action<T> action, CancellationToken token)
        {
            ThrowArgumentNullException(source, nameof(source));
            ThrowArgumentNullException(action, nameof(action));

            foreach (var element in source)
            {
                ThrowOperationCanceledException(token.IsCancellationRequested);
                action(element);
            }
        }

        public static void ForEach<T>(this IEnumerable<T> source, Action<T, int> action) =>
            source.ForEach(action, CancellationToken.None);

        public static void ForEach<T>(this IEnumerable<T> source, Action<T, int> action, CancellationToken token)
        {
            ThrowArgumentNullException(source, nameof(source));
            ThrowArgumentNullException(action, nameof(action));

            var index = 0;

            foreach (var element in source)
            {
                ThrowOperationCanceledException(token.IsCancellationRequested);
                action(element, index++);
            }
        }

        private static void ThrowArgumentNullException<T>(T param, string paramName)
        {
            if (param == null)
                throw new ArgumentNullException(paramName);
        }

        private static void ThrowOperationCanceledException(bool isCanceled)
        {
            if (isCanceled)
                throw new OperationCanceledException();
        }
    }
}
