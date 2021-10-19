using System;
using Confluent.Kafka;


namespace ConfluentKafka
{
    class Program
    {
        static void Main(string[] args)
        {
            var p = new Producer();
            p.Fire();
            Console.WriteLine("Hello World!");
        }
    }

    class Producer
    {

        IProducer<Null, string> producer = null;
        public Producer()
        {
            var conf = new ProducerConfig { BootstrapServers = "localhost:9092" };
            producer = new ProducerBuilder<Null, string>(conf).Build();
        }
        void handlerFunc2(DeliveryReport<Null, string> r)
        {

        }

        public void Fire()
        {
            Action<DeliveryReport<Null, string>> handlerFunc1 = r =>
                Console.WriteLine(!r.Error.IsError
                    ? $"Delivered message to {r.TopicPartitionOffset}"
                    : $"Delivery Error: {r.Error.Reason}");

            for (int i = 0; i < 44; ++i)
            {
                producer.Produce("my-topic", new Message<Null, string> { Value = i.ToString() }, handlerFunc1);
            }

            // wait for up to 10 seconds for any inflight messages to be delivered.
            producer.Flush(TimeSpan.FromSeconds(10));

        }
    }

}



