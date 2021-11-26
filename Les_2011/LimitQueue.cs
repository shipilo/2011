using System.Collections;

namespace Les_2011
{
    internal class LimitQueue<T> : Queue
    {
        public override void Enqueue(object obj)
        {
            base.Enqueue(obj);
            if (Count == 4)
            {           
                Dequeue();                
            }
        }
    }
}
