// File: Core/Help.Container.cs
using System;
using System.Collections.Generic;

namespace BetaTestSupp.Core
{
    // Mini-DI container
    public sealed class HelpContainer
    {
        private readonly Dictionary<Type, object> _map = new();

        public HelpContainer Register<T>(T instance)
        {
            _map[typeof(T)] = instance!;
            return this;
        }
        public T Resolve<T>()
        {
            if (_map.TryGetValue(typeof(T), out var obj)) return (T)obj;
            throw new InvalidOperationException($"Servizio non registrato: {typeof(T).Name}");
        }
    }
}
