import React, { useState } from 'react';

interface Card {
  name: string;
  set: string;
  rarity: string;
  quantity: number;
}

const CollectionManager: React.FC = () => {
  const [cards, setCards] = useState<Card[]>([]);
  const [name, setName] = useState<string>('');
  const [set, setSet] = useState<string>('');
  const [rarity, setRarity] = useState<string>('');
  const [quantity, setQuantity] = useState<string>('');

  const addCard = async () => {
    const newCard: Card = { name, set, rarity, quantity: parseInt(quantity) };
    setCards([...cards, newCard]);
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let table = sheet.tables.getItemOrNullObject('CardCollection');
      await context.sync();
      if (table.isNullObject) {
        table = sheet.tables.add('A1:D1', true);
        table.name = 'CardCollection';
        table.getHeaderRowRange().values = [['Name', 'Set', 'Rarity', 'Quantity']];
      }
      table.rows.add(null, [[name, set, rarity, quantity]]);
      await context.sync();
    });
    setName('');
    setSet('');
    setRarity('');
    setQuantity('');
  };

  return (
    <div>
      <h1>WoW TCG Collection Manager</h1>
      <input type="text" placeholder="Name" value={name} onChange={(e) => setName(e.target.value)} />
      <input type="text" placeholder="Set" value={set} onChange={(e) => setSet(e.target.value)} />
      <input type="text" placeholder="Rarity" value={rarity} onChange={(e) => setRarity(e.target.value)} />
      <input type="number" placeholder="Quantity" value={quantity} onChange={(e) => setQuantity(e.target.value)} />
      <button onClick={addCard}>Add Card</button>
      <h2>Collection</h2>
      <ul>
        {cards.map((card, index) => (
          <li key={index}>{`${card.name} - ${card.set} - ${card.rarity} - ${card.quantity}`}</li>
        ))}
      </ul>
    </div>
  );
};

export default CollectionManager;