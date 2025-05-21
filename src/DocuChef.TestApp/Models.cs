//namespace DocuChef.TestApp
//{


//    /// <summary>
//    /// 카테고리 정보를 나타내는 클래스
//    /// </summary>
//    public class Category
//    {
//        /// <summary>
//        /// 카테고리 ID
//        /// </summary>
//        public int Id { get; set; }

//        /// <summary>
//        /// 카테고리 이름
//        /// </summary>
//        public string Name { get; set; }

//        /// <summary>
//        /// 카테고리에 속한 제품 목록
//        /// </summary>
//        public List<Product> Products { get; private set; }

//        /// <summary>
//        /// 기본 생성자
//        /// </summary>
//        public Category()
//        {
//            Products = new List<Product>();
//        }

//        /// <summary>
//        /// ID와 이름으로 카테고리를 초기화하는 생성자
//        /// </summary>
//        public Category(int id, string name)
//        {
//            Id = id;
//            Name = name;
//            Products = new List<Product>();
//        }

//        /// <summary>
//        /// 카테고리에 제품을 추가합니다.
//        /// </summary>
//        public void AddProduct(Product product)
//        {
//            Products.Add(product);
//        }

//        /// <summary>
//        /// 카테고리에 대한 문자열 표현을 반환합니다.
//        /// </summary>
//        public override string ToString()
//        {
//            return $"{Name} (ID: {Id})";
//        }
//    }

//    /// <summary>
//    /// 제품 정보를 나타내는 클래스
//    /// </summary>
//    public class Product
//    {
//        /// <summary>
//        /// 제품 ID
//        /// </summary>
//        public int Id { get; set; }

//        /// <summary>
//        /// 제품 이름
//        /// </summary>
//        public string Name { get; set; }

//        /// <summary>
//        /// 제품 가격
//        /// </summary>
//        public decimal Price { get; set; }

//        /// <summary>
//        /// 제품 설명
//        /// </summary>
//        public string Description { get; set; }

//        /// <summary>
//        /// 기본 생성자
//        /// </summary>
//        public Product()
//        {
//        }

//        /// <summary>
//        /// 모든 속성을 초기화하는 생성자
//        /// </summary>
//        public Product(int id, string name, decimal price, string description)
//        {
//            Id = id;
//            Name = name;
//            Price = price;
//            Description = description;
//        }

//        /// <summary>
//        /// 제품에 대한 문자열 표현을 반환합니다.
//        /// </summary>
//        public override string ToString()
//        {
//            return $"{Name}: {Price:C} - {Description}";
//        }
//    }
//}