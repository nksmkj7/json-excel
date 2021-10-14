import { add }  from "../src/index";

describe("test add function", () => {
    it("should return 5 for add(2,3)", () => {
        expect(add(2,3)).toBe(5)
    });
})