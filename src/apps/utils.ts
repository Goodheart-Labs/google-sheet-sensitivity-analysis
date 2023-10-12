export const error = (message: string) => {
  console.error(message);

  return {
    success: false,
    message,
  };
};
