import * as React from 'react';
import { cn } from '@/src/lib/utils';

export interface InputProps extends React.InputHTMLAttributes<HTMLInputElement> {}

const Input = React.forwardRef<HTMLInputElement, InputProps>(
  ({ className, type, ...props }, ref) => {
    return (
      <input
        type={type}
        className={cn(
          'flex h-9 w-full rounded-md border border-white/15 bg-black px-3 py-1 text-sm text-white transition duration-200 ease-[cubic-bezier(0.32,0.72,0,1)] file:border-0 file:bg-transparent file:text-sm file:font-medium placeholder:text-white/35 focus-visible:border-white/40 focus-visible:outline-none disabled:cursor-not-allowed disabled:opacity-40',
          className
        )}
        ref={ref}
        {...props}
      />
    );
  }
);
Input.displayName = 'Input';

export { Input };
