using System;

namespace Invim.Restxcel.Helpers
{
    public class TryCatchHelper
    {
        public static T TryCatchIgnore<T>(Func<T> func)
        {
            try
            {
                return func();
            }
            catch
            {
                return default;
            }
        }

        public static void TryCatchIgnore(Action action)
        {
            try
            {
                action();
            }
            catch
            {
                
            }
        }
    }
}
