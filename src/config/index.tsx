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
    pessimisticInputColumnColumnIndex: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('baseInputColumnColumnIndex'),
          yup.ref('optimisticInputColumnColumnIndex'),
          yup.ref('baseOutputColumnColumnIndex'),
          yup.ref('optimisticOutputColumnColumnIndex'),
          yup.ref('pessimisticOutputColumnColumnIndex'),
          yup.ref('scenarioSwitcherColumnIndex'),
        ],
        'Pessimistic Scenario must not be in the same column as Base Scenario, Optimistic Scenario, or Scenario Switcher',
      ),
    baseInputColumnColumnIndex: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('pessimisticInputColumnColumnIndex'),
          yup.ref('optimisticInputColumnColumnIndex'),
          yup.ref('pessimisticOutputColumnColumnIndex'),
          yup.ref('optimisticOutputColumnColumnIndex'),
          yup.ref('baseOutputColumnColumnIndex'),
          yup.ref('scenarioSwitcherColumnIndex'),
        ],
        'Base Scenario must not be in the same column as Pessimistic Scenario, Optimistic Scenario, or Scenario Switcher',
      ),
    optimisticInputColumnColumnIndex: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('pessimisticInputColumnColumnIndex'),
          yup.ref('baseInputColumnColumnIndex'),
          yup.ref('pessimisticOutputColumnColumnIndex'),
          yup.ref('baseOutputColumnColumnIndex'),
          yup.ref('optimisticOutputColumnColumnIndex'),
          yup.ref('scenarioSwitcherColumnIndex'),
        ],
        'Optimistic Scenario must not be in the same column as Pessimistic Scenario, Base Scenario, or Scenario Switcher',
      ),
    pessimisticOutputColumnColumnIndex: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('baseInputColumnColumnIndex'),
          yup.ref('optimisticInputColumnColumnIndex'),
          yup.ref('pessimisticInputColumnColumnIndex'),
          yup.ref('baseOutputColumnColumnIndex'),
          yup.ref('optimisticOutputColumnColumnIndex'),
          yup.ref('scenarioSwitcherColumnIndex'),
        ],
        'Pessimistic Scenario must not be in the same column as Base Scenario, Optimistic Scenario, or Scenario Switcher',
      ),
    baseOutputColumnColumnIndex: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('pessimisticInputColumnColumnIndex'),
          yup.ref('optimisticInputColumnColumnIndex'),
          yup.ref('baseInputColumnColumnIndex'),
          yup.ref('pessimisticOutputColumnColumnIndex'),
          yup.ref('optimisticOutputColumnColumnIndex'),
          yup.ref('scenarioSwitcherColumnIndex'),
        ],
        'Base Scenario must not be in the same column as Pessimistic Scenario, Optimistic Scenario, or Scenario Switcher',
      ),
    optimisticOutputColumnColumnIndex: yup
      .string()
      .required()
      .matches(/^[A-Z]+$/)
      .notOneOf(
        [
          yup.ref('pessimisticInputColumnColumnIndex'),
          yup.ref('baseInputColumnColumnIndex'),
          yup.ref('optimisticInputColumnColumnIndex'),
          yup.ref('pessimisticOutputColumnColumnIndex'),
          yup.ref('baseOutputColumnColumnIndex'),
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
          baseInputColumnColumnIndex: string;
          pessimisticInputColumnColumnIndex: string;
          optimisticInputColumnColumnIndex: string;
          baseOutputColumnColumnIndex: string;
          pessimisticOutputColumnColumnIndex: string;
          optimisticOutputColumnColumnIndex: string;
        }) => {
          setValue(
            'scenarioSwitcherColumnIndex',
            config.scenarioSwitcherColumnIndex,
          );

          setValue('modelOutputCellIndex', config.modelOutputCellIndex);

          setValue(
            'baseInputColumnColumnIndex',
            config.baseInputColumnColumnIndex,
          );
          setValue(
            'pessimisticInputColumnColumnIndex',
            config.pessimisticInputColumnColumnIndex,
          );
          setValue(
            'optimisticInputColumnColumnIndex',
            config.optimisticInputColumnColumnIndex,
          );

          setValue(
            'baseOutputColumnColumnIndex',
            config.baseOutputColumnColumnIndex,
          );
          setValue(
            'pessimisticOutputColumnColumnIndex',
            config.pessimisticOutputColumnColumnIndex,
          );
          setValue(
            'optimisticOutputColumnColumnIndex',
            config.optimisticOutputColumnColumnIndex,
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

  const pessimisticInputColumnColumnIndex = watch(
    'pessimisticInputColumnColumnIndex',
  );
  const baseInputColumnColumnIndex = watch('baseInputColumnColumnIndex');
  const optimisticInputColumnColumnIndex = watch(
    'optimisticInputColumnColumnIndex',
  );
  const pessimisticOutputColumnColumnIndex = watch(
    'pessimisticOutputColumnColumnIndex',
  );
  const baseOutputColumnColumnIndex = watch('baseOutputColumnColumnIndex');
  const optimisticOutputColumnColumnIndex = watch(
    'optimisticOutputColumnColumnIndex',
  );

  // Inputs

  useEffect(() => {
    if (
      pessimisticInputColumnColumnIndex &&
      !baseInputColumnColumnIndex &&
      !optimisticInputColumnColumnIndex
    ) {
      setValue(
        'baseInputColumnColumnIndex',
        String.fromCharCode(
          pessimisticInputColumnColumnIndex.charCodeAt(0) + 1,
        ),
      );
      setValue(
        'optimisticInputColumnColumnIndex',
        String.fromCharCode(
          pessimisticInputColumnColumnIndex.charCodeAt(0) + 2,
        ),
      );
    }
  }, [
    pessimisticInputColumnColumnIndex,
    baseInputColumnColumnIndex,
    optimisticInputColumnColumnIndex,
    setValue,
  ]);

  // Outputs

  useEffect(() => {
    if (
      pessimisticOutputColumnColumnIndex &&
      !baseOutputColumnColumnIndex &&
      !optimisticOutputColumnColumnIndex
    ) {
      setValue(
        'baseOutputColumnColumnIndex',
        String.fromCharCode(
          pessimisticOutputColumnColumnIndex.charCodeAt(0) + 1,
        ),
      );
      setValue(
        'optimisticOutputColumnColumnIndex',
        String.fromCharCode(
          pessimisticOutputColumnColumnIndex.charCodeAt(0) + 2,
        ),
      );
    }
  }, [
    pessimisticOutputColumnColumnIndex,
    baseOutputColumnColumnIndex,
    optimisticOutputColumnColumnIndex,
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
            htmlFor="pessimisticInputColumnColumnIndex"
            className="block text-sm font-medium text-gray-600"
          >
            Pessimistic Scenario (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm disabled:bg-gray-50"
            disabled={loading}
            {...register('pessimisticInputColumnColumnIndex')}
          />
          <span className="text-red-500">
            {errors.pessimisticInputColumnColumnIndex?.message}
          </span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="baseInputColumnColumnIndex"
            className="block text-sm font-medium text-gray-600"
          >
            Base Scenario (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm disabled:bg-gray-50"
            disabled={loading}
            {...register('baseInputColumnColumnIndex')}
          />
          <span className="text-red-500">
            {errors.baseInputColumnColumnIndex?.message}
          </span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="optimisticInputColumnColumnIndex"
            className="block text-sm font-medium text-gray-600"
          >
            Optimistic Scenario (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm disabled:bg-gray-50"
            disabled={loading}
            {...register('optimisticInputColumnColumnIndex')}
          />
          <span className="text-red-500">
            {errors.optimisticInputColumnColumnIndex?.message}
          </span>
        </div>
      </fieldset>

      <fieldset className="mb-4 p-4 border rounded">
        <legend className="text-sm font-medium text-gray-600">Outputs</legend>

        <div className="mb-4">
          <label
            htmlFor="pessimisticOutputColumnColumnIndex"
            className="block text-sm font-medium text-gray-600"
          >
            Pessimistic Scenario (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm disabled:bg-gray-50"
            disabled={loading}
            {...register('pessimisticOutputColumnColumnIndex')}
          />
          <span className="text-red-500">
            {errors.pessimisticOutputColumnColumnIndex?.message}
          </span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="baseOutputColumnColumnIndex"
            className="block text-sm font-medium text-gray-600"
          >
            Base Scenario (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm disabled:bg-gray-50"
            disabled={loading}
            {...register('baseOutputColumnColumnIndex')}
          />
          <span className="text-red-500">
            {errors.baseOutputColumnColumnIndex?.message}
          </span>
        </div>

        <div className="mb-4">
          <label
            htmlFor="optimisticOutputColumnColumnIndex"
            className="block text-sm font-medium text-gray-600"
          >
            Optimistic Scenario (Column)
          </label>
          <input
            type="text"
            className="mt-1 py-1 px-2 w-full border rounded-md text-sm disabled:bg-gray-50"
            disabled={loading}
            {...register('optimisticOutputColumnColumnIndex')}
          />
          <span className="text-red-500">
            {errors.optimisticOutputColumnColumnIndex?.message}
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
