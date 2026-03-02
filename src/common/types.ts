export type SuccessResponse<T> = {
  success: true;
  data: T;
  problem?: never;
};

export type ErrorResponse = {
  data?: never;
  success: false;
  problem: string;
};

export type ActionResponse<T> = ErrorResponse | SuccessResponse<T>;
