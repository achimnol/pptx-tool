export interface FieldError {
  /** The dotted field path, such as "majorFont.latin", or "name". */
  key: string;
  message: string;
}

/** An error response of the API, following Litestar's error format. */
export class ApiError extends Error {
  readonly status: number;
  readonly fieldErrors: FieldError[];
  readonly extra: unknown;

  constructor(status: number, detail: string, extra?: unknown) {
    super(detail);
    this.name = 'ApiError';
    this.status = status;
    this.extra = extra;
    this.fieldErrors = Array.isArray(extra)
      ? extra.filter(
          (e): e is FieldError =>
            typeof e === 'object' &&
            e !== null &&
            typeof (e as FieldError).key === 'string' &&
            typeof (e as FieldError).message === 'string',
        )
      : [];
  }
}

/** Convert the error body of an API response into an ApiError. */
export function toApiError(response: Response, body: unknown): ApiError {
  if (typeof body === 'object' && body !== null && 'detail' in body) {
    const {detail, extra} = body as {detail: unknown; extra?: unknown};
    return new ApiError(response.status, String(detail), extra);
  }
  return new ApiError(
    response.status,
    `The request failed (${response.status} ${response.statusText}).`,
  );
}
