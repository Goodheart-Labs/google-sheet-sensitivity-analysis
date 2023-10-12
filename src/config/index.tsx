/* eslint-disable node/no-unpublished-import */

import React, { useEffect, useState } from 'react';
import { createRoot } from 'react-dom/client';
import { useForm } from 'react-hook-form';
import { yupResolver } from '@hookform/resolvers/yup';
import clsx from 'clsx';
import * as yup from 'yup';

declare const google: any;

const schema = yup
  .object({
    scenarioSwitcherColumnIndex: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/),
    modelOutputCellIndex: yup
      .string()
      .required()
      .matches(/^[A-Z]+\d+$/)
      .test(
        'is-not-in-same-column',
        'Model Output must not be in the same column as Scenario Switcher',
        function (value) {
          const scenarioSwitcherColumnIndexValue =
            this.parent.scenarioSwitcherColumnIndex;
          const modelOutputCellIndexColumn = value.match(/^[A-Z]+/)?.[0];
          return (
            modelOutputCellIndexColumn !== scenarioSwitcherColumnIndexValue
          );
        },
      ),
    pessimisticColumnColumnIndex: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('baseColumnColumnIndex'),
          yup.ref('optimisticColumnColumnIndex'),
          yup.ref('scenarioSwitcherColumnIndex'),
        ],
        'Pessimistic Scenario must not be in the same column as Base Scenario, Optimistic Scenario, or Scenario Switcher',
      ),
    baseColumnColumnIndex: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('pessimisticColumnColumnIndex'),
          yup.ref('optimisticColumnColumnIndex'),
          yup.ref('scenarioSwitcherColumnIndex'),
        ],
        'Base Scenario must not be in the same column as Pessimistic Scenario, Optimistic Scenario, or Scenario Switcher',
      ),
    optimisticColumnColumnIndex: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('pessimisticColumnColumnIndex'),
          yup.ref('baseColumnColumnIndex'),
          yup.ref('scenarioSwitcherColumnIndex'),
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

  // Get config values on load

  const [loading, setLoading] = useState(true);

  useEffect(() => {
    google.script.run
      .withSuccessHandler(
        (config: {
          scenarioSwitcherColumnIndex: string;
          modelOutputCellIndex: string;
          pessimisticColumnColumnIndex: string;
          baseColumnColumnIndex: string;
          optimisticColumnColumnIndex: string;
        }) => {
          setValue(
            'scenarioSwitcherColumnIndex',
            config.scenarioSwitcherColumnIndex,
          );
          setValue('modelOutputCellIndex', config.modelOutputCellIndex);
          setValue(
            'pessimisticColumnColumnIndex',
            config.pessimisticColumnColumnIndex,
          );
          setValue('baseColumnColumnIndex', config.baseColumnColumnIndex);
          setValue(
            'optimisticColumnColumnIndex',
            config.optimisticColumnColumnIndex,
          );
          setLoading(false);
        },
      )
      .withFailureHandler((err: any) => {
        // TODO: Show error message
        console.error(err);
        setLoading(false);
      })
      .getConfigValues();
  }, []);

  // Handle C/D/E column updates

  const pessimisticColumnColumnIndex = watch('pessimisticColumnColumnIndex');
  const baseColumnColumnIndex = watch('baseColumnColumnIndex');
  const optimisticColumnColumnIndex = watch('optimisticColumnColumnIndex');

  useEffect(() => {
    if (
      pessimisticColumnColumnIndex &&
      !baseColumnColumnIndex &&
      !optimisticColumnColumnIndex
    ) {
      setValue(
        'baseColumnColumnIndex',
        String.fromCharCode(pessimisticColumnColumnIndex.charCodeAt(0) + 1),
      );
      setValue(
        'optimisticColumnColumnIndex',
        String.fromCharCode(pessimisticColumnColumnIndex.charCodeAt(0) + 2),
      );
    }
  }, [
    pessimisticColumnColumnIndex,
    baseColumnColumnIndex,
    optimisticColumnColumnIndex,
    setValue,
  ]);

  // Callbacks

  const onSubmit = (config: any) => {
    google.script.run
      .withSuccessHandler(() => {
        google.script.host.close();
      })
      .withFailureHandler((err: any) => {
        // TODO: Show error message
        console.error(err);
      })
      .setConfigValues({ config });
  };

  // Render

  return (
    <form id="config" onSubmit={handleSubmit(onSubmit)}>
      <fieldset className="mb-4 p-4 border rounded">
        <legend className="text-sm font-medium text-gray-600">Scenarios</legend>
        <div className="mb-4">
          <label
            htmlFor="scenarioSwitcherColumnIndex"
            className="block text-sm font-medium text-gray-600"
          >
            Scenario Switcher (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm disabled:bg-gray-50"
            disabled={loading}
            {...register('scenarioSwitcherColumnIndex')}
          />
          <span className="text-red-500">
            {errors.scenarioSwitcherColumnIndex?.message}
          </span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="modelOutputCellIndex"
            className="block text-sm font-medium text-gray-600"
          >
            Model Output (Cell)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm disabled:bg-gray-50"
            disabled={loading}
            {...register('modelOutputCellIndex')}
          />
          <span className="text-red-500">
            {errors.modelOutputCellIndex?.message}
          </span>
        </div>
      </fieldset>

      <fieldset className="mb-4 p-4 border rounded">
        <legend className="text-sm font-medium text-gray-600">Inputs</legend>

        <div className="mb-4">
          <label
            htmlFor="pessimisticColumnColumnIndex"
            className="block text-sm font-medium text-gray-600"
          >
            Pessimistic Scenario (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm disabled:bg-gray-50"
            disabled={loading}
            {...register('pessimisticColumnColumnIndex')}
          />
          <span className="text-red-500">
            {errors.pessimisticColumnColumnIndex?.message}
          </span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="baseColumnColumnIndex"
            className="block text-sm font-medium text-gray-600"
          >
            Base Scenario (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm disabled:bg-gray-50"
            disabled={loading}
            {...register('baseColumnColumnIndex')}
          />
          <span className="text-red-500">
            {errors.baseColumnColumnIndex?.message}
          </span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="optimisticColumnColumnIndex"
            className="block text-sm font-medium text-gray-600"
          >
            Optimistic Scenario (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm disabled:bg-gray-50"
            disabled={loading}
            {...register('optimisticColumnColumnIndex')}
          />
          <span className="text-red-500">
            {errors.optimisticColumnColumnIndex?.message}
          </span>
        </div>
      </fieldset>

      <button
        type="submit"
        className={clsx(
          'mt-4 bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded disabled:opacity-50 disabled:cursor-not-allowed disabled:pointer-events-none',
        )}
        disabled={!isValid || loading}
      >
        Save
      </button>
    </form>
  );
};

createRoot(document.getElementById('root') as any).render(<App />);
