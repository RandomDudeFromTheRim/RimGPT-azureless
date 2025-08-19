using System;
using Verse;

namespace RimGPT
{
	// for scheduling a recurring action
	// - updateIntervalFunc is a function that returns the time in ticks between invoking the action
	// - action is executed at each scheduled interval
	// - startImmediately to indicate the action will be executed as soon as the task starts
	//
        public struct UpdateTask
        {
                public int updateTickCounter;
                public Func<int> updateIntervalFunc;
                public Action<Map> action;

                public UpdateTask(Func<int> updateIntervalFunc, Action<Map> action, bool startImmediately)
                {
                        this.updateIntervalFunc = updateIntervalFunc;
                        this.action = action;

                        var interval = Math.Max(1, updateIntervalFunc());
                        updateTickCounter = startImmediately ? 0 : Rand.Range(interval / 2, interval);
                }
        }
}