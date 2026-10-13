import { describe, expect, it } from 'vitest'
import { cutRanges, keepRanges, remap, toSrt, toTxt } from './cuts.js'

const words = [
  { text: 'uno', start: 0, end: 1 },
  { text: 'dos', start: 1.5, end: 2 },
  { text: 'tres.', start: 3, end: 4 },
  { text: 'cuatro', start: 4, end: 5 },
]

describe('cuts', () => {
  it('une palabras borradas contiguas hasta la siguiente conservada', () => {
    expect(cutRanges(words, new Set([1, 2]), 5)).toEqual([[1.5, 4]])
  })
  it('borrar la última corta hasta el final', () => {
    expect(cutRanges(words, new Set([3]), 6)).toEqual([[4, 6]])
  })
  it('borrar la primera corta desde su inicio', () => {
    expect(cutRanges(words, new Set([0]), 5)).toEqual([[0, 1.5]])
  })
  it('sin borrados no hay cortes y se conserva todo', () => {
    expect(cutRanges(words, new Set(), 5)).toEqual([])
    expect(keepRanges([], 5)).toEqual([[0, 5]])
  })
  it('keepRanges es el complemento', () => {
    expect(keepRanges([[1.5, 4]], 5)).toEqual([[0, 1.5], [4, 5]])
    expect(keepRanges([[0, 1.5]], 5)).toEqual([[1.5, 5]])
    expect(keepRanges([[0, 5]], 5)).toEqual([])
  })
  it('remap descuenta lo cortado', () => {
    const cuts = [[1.5, 4]]
    expect(remap(1, cuts)).toBe(1)
    expect(remap(2, cuts)).toBe(1.5)
    expect(remap(4.5, cuts)).toBe(2)
  })
  it('srt y txt reflejan la edición', () => {
    const del = new Set([1, 2])
    const cuts = cutRanges(words, del, 5)
    expect(toTxt(words, del)).toBe('uno cuatro')
    expect(toSrt(words, del, cuts)).toBe('1\n00:00:00,000 --> 00:00:02,500\nuno cuatro\n')
  })
})
