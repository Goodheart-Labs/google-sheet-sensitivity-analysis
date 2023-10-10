/* eslint-disable node/no-unpublished-import */

import React from 'react';
import { createRoot } from 'react-dom/client';

const App = () => {
  return (
    <form id="config">
      <fieldset className="mb-4 p-4 border rounded">
        <legend className="text-sm font-medium text-gray-600">Scenarios</legend>
        <div className="mb-4">
          <label
            htmlFor="scenarioSwitcher"
            className="block text-sm font-medium text-gray-600"
          >
            Scenario Switcher (Column)
          </label>
          <input
            id="scenarioSwitcher"
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm"
            required
            pattern="^[A-Z]+$"
          />
          <span
            id="scenarioSwitcherError"
            className="text-red-500 hidden"
          ></span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="modelOutput"
            className="block text-sm font-medium text-gray-600"
          >
            Model Output (Cell)
          </label>
          <input
            id="modelOutput"
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm"
            required
            pattern="^[A-Z]+\d+$"
          />
          <span id="modelOutputError" className="text-red-500 hidden"></span>
        </div>
      </fieldset>

      <fieldset className="mb-4 p-4 border rounded">
        <legend className="text-sm font-medium text-gray-600">Inputs</legend>

        <div className="mb-4">
          <label
            htmlFor="pessimisticColumn"
            className="block text-sm font-medium text-gray-600"
          >
            Pessimistic Scenario (Column)
          </label>
          <input
            id="pessimisticColumn"
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm"
            required
            pattern="^[A-Z]+$"
          />
          <span
            id="pessimisticColumnError"
            className="text-red-500 hidden"
          ></span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="baseColumn"
            className="block text-sm font-medium text-gray-600"
          >
            Base Scenario (Column)
          </label>
          <input
            id="baseColumn"
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm"
            required
            pattern="^[A-Z]+$"
          />
          <span id="baseColumnError" className="text-red-500 hidden"></span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="optimisticColumn"
            className="block text-sm font-medium text-gray-600"
          >
            Optimistic Scenario (Column)
          </label>
          <input
            id="optimisticColumn"
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm"
            required
            pattern="^[A-Z]+$"
          />
          <span
            id="optimisticColumnError"
            className="text-red-500 hidden"
          ></span>
        </div>
      </fieldset>

      <button
        id="submit"
        type="submit"
        className="mt-4 bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded opacity-50 cursor-not-allowed pointer-events-none"
        disabled
      >
        Save
      </button>
    </form>
  );
};

createRoot(document.getElementById('root') as any).render(<App />);
