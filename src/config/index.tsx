/* eslint-disable node/no-unpublished-import */

import React, { useEffect } from 'react';
import { createRoot } from 'react-dom/client';
import { useForm } from 'react-hook-form';
import { yupResolver } from '@hookform/resolvers/yup';
import clsx from 'clsx';
import * as yup from 'yup';

declare const google: any;

const schema = yup
  .object({
    scenarioSwitcher: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/),
    modelOutput: yup
      .string()
      .required()
      .matches(/^[A-Z]+\d+$/)
      .test(
        'is-not-in-same-column',
        'Model Output must not be in the same column as Scenario Switcher',
        function (value) {
          const scenarioSwitcherValue = this.parent.scenarioSwitcher;
          const modelOutputColumn = value.match(/^[A-Z]+/)?.[0];
          return modelOutputColumn !== scenarioSwitcherValue;
        },
      ),
    pessimisticColumn: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('baseColumn'),
          yup.ref('optimisticColumn'),
          yup.ref('scenarioSwitcher'),
        ],
        'Pessimistic Scenario must not be in the same column as Base Scenario, Optimistic Scenario, or Scenario Switcher',
      ),
    baseColumn: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('pessimisticColumn'),
          yup.ref('optimisticColumn'),
          yup.ref('scenarioSwitcher'),
        ],
        'Base Scenario must not be in the same column as Pessimistic Scenario, Optimistic Scenario, or Scenario Switcher',
      ),
    optimisticColumn: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('pessimisticColumn'),
          yup.ref('baseColumn'),
          yup.ref('scenarioSwitcher'),
        ],
        'Optimistic Scenario must not be in the same column as Pessimistic Scenario, Base Scenario, or Scenario Switcher',
      ),
  })
  .required();

const App = () => {
  const {
    register,
    handleSubmit,
    watch,
    setValue,
    formState: { errors, isValid },
  } = useForm({
    mode: 'onBlur',
    resolver: yupResolver(schema),
  });

  const pessimisticColumn = watch('pessimisticColumn');
  const baseColumn = watch('baseColumn');
  const optimisticColumn = watch('optimisticColumn');

  useEffect(() => {
    if (pessimisticColumn && !baseColumn && !optimisticColumn) {
      setValue(
        'baseColumn',
        String.fromCharCode(pessimisticColumn.charCodeAt(0) + 1),
      );
      setValue(
        'optimisticColumn',
        String.fromCharCode(pessimisticColumn.charCodeAt(0) + 2),
      );
    }
  }, [pessimisticColumn, baseColumn, optimisticColumn, setValue]);

  const onSubmit = (data: any) => {
    google.script.run
      .withSuccessHandler(() => {
        google.script.host.close();
      })
      .withFailureHandler((err: any) => {
        // TODO: Show error message
        console.error(err);
      })
      .processJSONData({ data });
  };

  return (
    <form id="config" onSubmit={handleSubmit(onSubmit)}>
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
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm"
            {...register('scenarioSwitcher')}
          />
          <span className="text-red-500">
            {errors.scenarioSwitcher?.message}
          </span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="modelOutput"
            className="block text-sm font-medium text-gray-600"
          >
            Model Output (Cell)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm"
            {...register('modelOutput')}
          />
          <span className="text-red-500">{errors.modelOutput?.message}</span>
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
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm"
            {...register('pessimisticColumn')}
          />
          <span className="text-red-500">
            {errors.pessimisticColumn?.message}
          </span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="baseColumn"
            className="block text-sm font-medium text-gray-600"
          >
            Base Scenario (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm"
            {...register('baseColumn')}
          />
          <span className="text-red-500">{errors.baseColumn?.message}</span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="optimisticColumn"
            className="block text-sm font-medium text-gray-600"
          >
            Optimistic Scenario (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm"
            {...register('optimisticColumn')}
          />
          <span className="text-red-500">
            {errors.optimisticColumn?.message}
          </span>
        </div>
      </fieldset>

      <button
        type="submit"
        className={clsx(
          'mt-4 bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded',
          !isValid && 'opacity-50 cursor-not-allowed pointer-events-none',
        )}
        disabled={!isValid}
      >
        Save
      </button>
    </form>
  );
};

createRoot(document.getElementById('root') as any).render(<App />);
