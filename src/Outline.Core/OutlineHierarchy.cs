using System;
using System.Collections.Generic;
using System.Linq;

namespace Outline.Core
{
    /// <summary>Compress absent outline depths without changing the nearest enclosing row tree.</summary>
    public static class OutlineHierarchy
    {
        public static Dictionary<int,int> Compact(IDictionary<int,int> source)
        {
            if(source==null)throw new ArgumentNullException(nameof(source));
            var result=new Dictionary<int,int>();var stack=new List<int>();
            foreach(var item in source.OrderBy(x=>x.Key))
            {
                if(item.Value<0||item.Value>7){result[item.Key]=item.Value;stack.Clear();continue;}
                while(stack.Count>0&&stack[stack.Count-1]>=item.Value)stack.RemoveAt(stack.Count-1);
                result[item.Key]=stack.Count;
                stack.Add(item.Value);
            }
            return result;
        }
    }
}
