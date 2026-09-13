export class EventSource {
  #listeners = new Set();
  Subscribe(listener) {
    if (typeof listener !== 'function') throw new TypeError('Expected a listener');
    this.#listeners.add(listener); let active = true;
    const dispose = () => { if (active) { active = false; this.#listeners.delete(listener); } };
    return { Dispose: dispose, dispose, unsubscribe: dispose };
  }
  subscribe(listener) { return this.Subscribe(listener); }
  Add(listener) { return this.Subscribe(listener); }
  Emit(value) { for (const listener of [...this.#listeners]) { try { listener(value); } catch (e) { console.error('GridWeb event listener failed', e); } } }
  Clear() { this.#listeners.clear(); }
  get Count() { return this.#listeners.size; }
}
export class ObservableObject {
  constructor() { this.PropertyChanged = new EventSource(); }
  SetProperty(name, value) {
    if (['__proto__', 'constructor', 'prototype'].includes(name)) throw new TypeError('Unsafe property');
    const oldValue = this[name]; if (Object.is(oldValue, value)) return false;
    this[name] = value; this.PropertyChanged.Emit({ Sender: this, PropertyName: name, OldValue: oldValue, NewValue: value }); return true;
  }
  Dispose() { this.PropertyChanged.Clear(); }
}
export class RelayCommand {
  constructor(execute, canExecute = () => true) { this._execute = execute; this._canExecute = canExecute; this.CanExecuteChanged = new EventSource(); }
  CanExecute(parameter) { return !!this._canExecute(parameter); }
  Execute(parameter) { if (this.CanExecute(parameter)) return this._execute(parameter); }
  NotifyCanExecuteChanged() { this.CanExecuteChanged.Emit(this); }
  Dispose() { this.CanExecuteChanged.Clear(); }
}
export class ObservableCollection {
  constructor(items = []) { this._items = [...items]; this.CollectionChanged = new EventSource(); }
  get Count() { return this._items.length; }
  Get(index) { return this._items[index]; }
  Add(item) { this._items.push(item); this.CollectionChanged.Emit({ Action: 'Add', NewItems: [item], Index: this.Count - 1 }); return item; }
  Remove(item) { const i = this._items.indexOf(item); if (i < 0) return false; this._items.splice(i, 1); this.CollectionChanged.Emit({ Action: 'Remove', OldItems: [item], Index: i }); return true; }
  [Symbol.iterator]() { return this._items[Symbol.iterator](); }
}
export function bind(source, path, target, property, { mode = 'OneWay', event = 'change', convert = v => v } = {}) {
  const keys = path.split('.'); if (keys.some(x => ['__proto__', 'constructor', 'prototype'].includes(x))) throw new TypeError('Unsafe binding path');
  const read = () => keys.reduce((a, b) => a?.[b], source);
  const update = () => { target[property] = convert(read()); }; update();
  const subscription = source.PropertyChanged?.Subscribe(update);
  const write = () => { const owner = keys.slice(0, -1).reduce((a, b) => a[b], source), key = keys.at(-1); if (owner.SetProperty) owner.SetProperty(key, target[property]); else owner[key] = target[property]; };
  if (mode === 'TwoWay') target.addEventListener(event, write);
  return { Dispose() { subscription?.Dispose(); target.removeEventListener(event, write); } };
}
