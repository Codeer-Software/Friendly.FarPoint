using System;
using System.Linq;

namespace Friendly.FarPoint.Generator
{
    public static class EventAdapter
    {
        public static Action Add(object target, string eventName, Action<object, dynamic> handler)
        {
            var eventIinfo = target.GetType().GetEvent(eventName);

            var handlerType = eventIinfo.EventHandlerType;
            var arguments = handlerType.GetMethod("Invoke").GetParameters().Select(e => e.ParameterType).ToArray();
            var conType = typeof(Adapter<,>).MakeGenericType(arguments[0], arguments[1]);
            var eh = conType.GetMethod("Adapt").Invoke(null, new object[] { handlerType, handler }) as Delegate;

            eventIinfo.AddEventHandler(target, eh);
            return () => eventIinfo.RemoveEventHandler(target, eh);
        }

        public static class Adapter<O, E>
        {
            public class Core
            {
                Action<object, dynamic> _core;
                public Core(Action<object, dynamic> core) { _core = core; }
                public void Invoke(O o, E e) => _core(o, e);
            }

            public static Delegate Adapt(Type handlerType, Action<object, dynamic> handler)
            {
                var core = new Core(handler);
                return Delegate.CreateDelegate(handlerType, core, core.GetType().GetMethod("Invoke"));
            }
        }
    }
}
